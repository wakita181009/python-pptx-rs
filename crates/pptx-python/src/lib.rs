//! PyO3 bindings exposing a python-pptx-compatible API.
//!
//! Every proxy object holds `Py<Presentation>` plus stable node ids into the
//! owning part's DOM, so one attribute access is a single Rust call instead of
//! python-pptx's descriptor/lxml call chain.

use std::collections::HashMap;
use std::path::PathBuf;

use pyo3::exceptions::{PyIndexError, PyKeyError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyIterator, PyList};

use pptx_core::dom::{Document, NodeId};
use pptx_core::error::Error;
use pptx_core::ns;
use pptx_core::opc::Package;

fn to_py(e: Error) -> PyErr {
    match e {
        Error::Io(e) => e.into(),
        Error::PartNotFound(p) => PyKeyError::new_err(p),
        other => PyValueError::new_err(other.to_string()),
    }
}

const OFFICE_DOCUMENT_RELTYPE_SUFFIX: &str = "/officeDocument";
/// python-pptx maps `a:br` to a vertical-tab character in text getters/setters.
const VERTICAL_TAB: char = '\u{0B}';
const SHAPE_TAGS: [&str; 5] = ["sp", "pic", "graphicFrame", "grpSp", "cxnSp"];

// ---------------------------------------------------------------------------
// Presentation
// ---------------------------------------------------------------------------

#[pyclass(module = "pptx_rs._core")]
pub struct Presentation {
    pkg: Package,
    part_name: String,
    /// (slide part name, slide id) in presentation order.
    slide_entries: Vec<(String, u32)>,
    /// Cached next drawing-object id per slide part (O(1) shape adds).
    next_shape_ids: HashMap<String, u32>,
}

impl Presentation {
    fn load(mut pkg: Package) -> PyResult<Self> {
        let root_rels = pkg.rels("").map_err(to_py)?;
        let pres_part = root_rels
            .iter()
            .find(|r| r.reltype.ends_with(OFFICE_DOCUMENT_RELTYPE_SUFFIX))
            .map(|r| pkg.resolve_target("", &r.target))
            .ok_or_else(|| PyValueError::new_err("package is not a PresentationML package"))?;

        let rels = pkg.rels(&pres_part).map_err(to_py)?;
        let rel_targets: HashMap<String, String> = rels
            .into_iter()
            .map(|r| {
                let target = pkg.resolve_target(&pres_part, &r.target);
                (r.id, target)
            })
            .collect();

        let doc = pkg.doc(&pres_part).map_err(to_py)?;
        let mut slide_entries = Vec::new();
        if let Some(sld_id_lst) = doc.first_child_named(doc.root, ns::P, "sldIdLst") {
            for sld_id in doc.children_named(sld_id_lst, ns::P, "sldId") {
                let rid = doc.attr(sld_id, "r:id").unwrap_or_default();
                let id: u32 = doc.attr(sld_id, "id").unwrap_or("0").parse().unwrap_or(0);
                if let Some(part) = rel_targets.get(rid) {
                    slide_entries.push((part.clone(), id));
                }
            }
        }

        Ok(Presentation {
            pkg,
            part_name: pres_part,
            slide_entries,
            next_shape_ids: HashMap::new(),
        })
    }

    fn doc(&mut self, part: &str) -> PyResult<&Document> {
        self.pkg.doc(part).map_err(to_py)
    }

    fn doc_mut(&mut self, part: &str) -> PyResult<&mut Document> {
        self.pkg.doc_mut(part).map_err(to_py)
    }

    /// Next available drawing-object id in `part`, computed once then O(1).
    fn take_next_shape_id(&mut self, part: &str) -> PyResult<u32> {
        if !self.next_shape_ids.contains_key(part) {
            let doc = self.pkg.doc(part).map_err(to_py)?;
            let mut max_id = 1u32;
            doc.walk(doc.root, &mut |d, id| {
                if d.is(id, ns::P, "cNvPr")
                    && let Some(n) = d.attr(id, "id").and_then(|v| v.parse::<u32>().ok())
                {
                    max_id = max_id.max(n);
                }
            });
            self.next_shape_ids.insert(part.to_string(), max_id + 1);
        }
        let next = self.next_shape_ids.get_mut(part).expect("just inserted");
        let id = *next;
        *next += 1;
        Ok(id)
    }

    fn sld_sz(&mut self, attr: &str) -> PyResult<Option<i64>> {
        let part = self.part_name.clone();
        let doc = self.doc(&part)?;
        Ok(doc
            .first_child_named(doc.root, ns::P, "sldSz")
            .and_then(|n| doc.attr(n, attr))
            .and_then(|v| v.parse().ok()))
    }
}

#[pymethods]
impl Presentation {
    #[staticmethod]
    fn from_bytes(data: Vec<u8>) -> PyResult<Self> {
        Self::load(Package::from_bytes(&data).map_err(to_py)?)
    }

    #[staticmethod]
    fn from_path(path: PathBuf) -> PyResult<Self> {
        let data = std::fs::read(path)?;
        Self::load(Package::from_bytes(&data).map_err(to_py)?)
    }

    #[getter]
    fn slides(slf: Py<Self>) -> Slides {
        Slides { prs: slf }
    }

    #[getter]
    fn slide_width(&mut self) -> PyResult<Option<i64>> {
        self.sld_sz("cx")
    }

    #[getter]
    fn slide_height(&mut self) -> PyResult<Option<i64>> {
        self.sld_sz("cy")
    }

    fn save(&self, target: &Bound<'_, PyAny>) -> PyResult<()> {
        let bytes = self.pkg.save_to_bytes().map_err(to_py)?;
        if let Ok(path) = target.extract::<PathBuf>() {
            std::fs::write(path, bytes)?;
        } else {
            target.call_method1("write", (PyBytes::new(target.py(), &bytes),))?;
        }
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

#[pyclass(module = "pptx_rs._core")]
pub struct Slides {
    prs: Py<Presentation>,
}

impl Slides {
    fn entries(&self, py: Python<'_>) -> Vec<(String, u32)> {
        self.prs.borrow(py).slide_entries.clone()
    }
}

#[pymethods]
impl Slides {
    fn __len__(&self, py: Python<'_>) -> usize {
        self.prs.borrow(py).slide_entries.len()
    }

    fn __getitem__(&self, py: Python<'_>, idx: isize) -> PyResult<Slide> {
        let entries = self.entries(py);
        let len = entries.len() as isize;
        let i = if idx < 0 { idx + len } else { idx };
        if i < 0 || i >= len {
            return Err(PyIndexError::new_err("slide index out of range"));
        }
        let (part, slide_id) = entries[i as usize].clone();
        Ok(Slide {
            prs: self.prs.clone_ref(py),
            part,
            slide_id,
        })
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyIterator>> {
        let slides: Vec<Slide> = self
            .entries(py)
            .into_iter()
            .map(|(part, slide_id)| Slide {
                prs: self.prs.clone_ref(py),
                part,
                slide_id,
            })
            .collect();
        Ok(PyList::new(py, slides)?.try_iter()?.unbind())
    }
}

// ---------------------------------------------------------------------------
// Slide
// ---------------------------------------------------------------------------

#[pyclass(module = "pptx_rs._core")]
pub struct Slide {
    prs: Py<Presentation>,
    part: String,
    slide_id: u32,
}

#[pymethods]
impl Slide {
    #[getter]
    fn slide_id(&self) -> u32 {
        self.slide_id
    }

    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<String> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(doc
            .first_child_named(doc.root, ns::P, "cSld")
            .and_then(|n| doc.attr(n, "name"))
            .unwrap_or_default()
            .to_string())
    }

    #[getter]
    fn shapes(&self, py: Python<'_>) -> SlideShapes {
        SlideShapes {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
        }
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .extract::<PyRef<'_, Slide>>()
            .is_ok_and(|o| self.prs.as_ptr() == o.prs.as_ptr() && self.part == o.part)
    }
}

// ---------------------------------------------------------------------------
// SlideShapes
// ---------------------------------------------------------------------------

#[pyclass(module = "pptx_rs._core")]
pub struct SlideShapes {
    prs: Py<Presentation>,
    part: String,
}

fn sp_tree(doc: &Document) -> PyResult<NodeId> {
    doc.first_child_named(doc.root, ns::P, "cSld")
        .and_then(|c| doc.first_child_named(c, ns::P, "spTree"))
        .ok_or_else(|| PyValueError::new_err("slide has no p:spTree"))
}

fn shape_nodes(doc: &Document) -> PyResult<Vec<NodeId>> {
    let tree = sp_tree(doc)?;
    Ok(doc
        .child_elements(tree)
        .into_iter()
        .filter(|&n| {
            let el = doc.get(n);
            el.ns.as_deref() == Some(ns::P) && SHAPE_TAGS.contains(&el.local.as_str())
        })
        .collect())
}

impl SlideShapes {
    fn nodes(&self, py: Python<'_>) -> PyResult<Vec<NodeId>> {
        let mut prs = self.prs.borrow_mut(py);
        shape_nodes(prs.doc(&self.part)?)
    }

    fn shape(&self, py: Python<'_>, node: NodeId) -> Shape {
        Shape {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node,
        }
    }
}

#[pymethods]
impl SlideShapes {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.nodes(py)?.len())
    }

    fn __getitem__(&self, py: Python<'_>, idx: isize) -> PyResult<Shape> {
        let nodes = self.nodes(py)?;
        let len = nodes.len() as isize;
        let i = if idx < 0 { idx + len } else { idx };
        if i < 0 || i >= len {
            return Err(PyIndexError::new_err("shape index out of range"));
        }
        Ok(self.shape(py, nodes[i as usize]))
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyIterator>> {
        let shapes: Vec<Shape> = self
            .nodes(py)?
            .into_iter()
            .map(|n| self.shape(py, n))
            .collect();
        Ok(PyList::new(py, shapes)?.try_iter()?.unbind())
    }

    fn add_textbox(
        &self,
        py: Python<'_>,
        left: i64,
        top: i64,
        width: i64,
        height: i64,
    ) -> PyResult<Shape> {
        let mut prs = self.prs.borrow_mut(py);
        let shape_id = prs.take_next_shape_id(&self.part)?;
        let name = format!("TextBox {}", shape_id - 1);
        let doc = prs.doc_mut(&self.part)?;
        let tree = sp_tree(doc)?;
        let sp = new_textbox_sp(doc, shape_id, &name, left, top, width, height);
        doc.append_child(tree, sp);
        drop(prs);
        Ok(self.shape(py, sp))
    }
}

/// Build a `p:sp` textbox subtree matching python-pptx's textbox template.
fn new_textbox_sp(
    doc: &mut Document,
    shape_id: u32,
    name: &str,
    left: i64,
    top: i64,
    width: i64,
    height: i64,
) -> NodeId {
    let id_s = shape_id.to_string();
    let (x, y) = (left.to_string(), top.to_string());
    let (cx, cy) = (width.to_string(), height.to_string());

    let sp = doc.create_element(ns::P, "p", "sp", &[]);

    let nv_sp_pr = doc.create_element(ns::P, "p", "nvSpPr", &[]);
    let c_nv_pr = doc.create_element(
        ns::P,
        "p",
        "cNvPr",
        &[("id", id_s.as_str()), ("name", name)],
    );
    let c_nv_sp_pr = doc.create_element(ns::P, "p", "cNvSpPr", &[("txBox", "1")]);
    let nv_pr = doc.create_element(ns::P, "p", "nvPr", &[]);
    doc.append_child(nv_sp_pr, c_nv_pr);
    doc.append_child(nv_sp_pr, c_nv_sp_pr);
    doc.append_child(nv_sp_pr, nv_pr);
    doc.append_child(sp, nv_sp_pr);

    let sp_pr = doc.create_element(ns::P, "p", "spPr", &[]);
    let xfrm = doc.create_element(ns::A, "a", "xfrm", &[]);
    let off = doc.create_element(ns::A, "a", "off", &[("x", x.as_str()), ("y", y.as_str())]);
    let ext = doc.create_element(
        ns::A,
        "a",
        "ext",
        &[("cx", cx.as_str()), ("cy", cy.as_str())],
    );
    doc.append_child(xfrm, off);
    doc.append_child(xfrm, ext);
    doc.append_child(sp_pr, xfrm);
    let prst_geom = doc.create_element(ns::A, "a", "prstGeom", &[("prst", "rect")]);
    let av_lst = doc.create_element(ns::A, "a", "avLst", &[]);
    doc.append_child(prst_geom, av_lst);
    doc.append_child(sp_pr, prst_geom);
    let no_fill = doc.create_element(ns::A, "a", "noFill", &[]);
    doc.append_child(sp_pr, no_fill);
    doc.append_child(sp, sp_pr);

    let tx_body = doc.create_element(ns::P, "p", "txBody", &[]);
    let body_pr = doc.create_element(ns::A, "a", "bodyPr", &[("wrap", "none")]);
    let sp_auto_fit = doc.create_element(ns::A, "a", "spAutoFit", &[]);
    doc.append_child(body_pr, sp_auto_fit);
    doc.append_child(tx_body, body_pr);
    let lst_style = doc.create_element(ns::A, "a", "lstStyle", &[]);
    doc.append_child(tx_body, lst_style);
    let p = doc.create_element(ns::A, "a", "p", &[]);
    doc.append_child(tx_body, p);
    doc.append_child(sp, tx_body);

    sp
}

// ---------------------------------------------------------------------------
// Shape
// ---------------------------------------------------------------------------

#[pyclass(module = "pptx_rs._core")]
pub struct Shape {
    prs: Py<Presentation>,
    part: String,
    node: NodeId,
}

fn c_nv_pr(doc: &Document, shape: NodeId) -> Option<NodeId> {
    doc.child_elements(shape)
        .into_iter()
        .find_map(|child| doc.first_child_named(child, ns::P, "cNvPr"))
}

fn tx_body(doc: &Document, shape: NodeId) -> Option<NodeId> {
    doc.first_child_named(shape, ns::P, "txBody")
}

/// The element holding `a:xfrm` (or `p:xfrm`), by shape kind:
/// `p:sp`/`p:pic`/`p:cxnSp` → `p:spPr`, `p:grpSp` → `p:grpSpPr`,
/// `p:graphicFrame` → the frame element itself (direct `p:xfrm` child).
fn xfrm_parent(doc: &Document, shape: NodeId) -> Option<(NodeId, &'static str)> {
    if doc.get(shape).local == "graphicFrame" {
        return Some((shape, ns::P));
    }
    if let Some(sp_pr) = doc.first_child_named(shape, ns::P, "spPr") {
        return Some((sp_pr, ns::A));
    }
    doc.first_child_named(shape, ns::P, "grpSpPr")
        .map(|n| (n, ns::A))
}

impl Shape {
    fn xfrm_val(&self, py: Python<'_>, child: &str, attr: &str) -> PyResult<Option<i64>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(xfrm_parent(doc, self.node)
            .and_then(|(parent, xfrm_ns)| doc.first_child_named(parent, xfrm_ns, "xfrm"))
            .and_then(|xfrm| doc.first_child_named(xfrm, ns::A, child))
            .and_then(|n| doc.attr(n, attr))
            .and_then(|v| v.parse().ok()))
    }

    fn set_xfrm_val(&self, py: Python<'_>, child: &str, attr: &str, value: i64) -> PyResult<()> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc_mut(&self.part)?;
        let (parent, xfrm_ns) = xfrm_parent(doc, self.node)
            .ok_or_else(|| PyValueError::new_err("shape has no transform parent element"))?;
        let xfrm = match doc.first_child_named(parent, xfrm_ns, "xfrm") {
            Some(x) => x,
            None => {
                let prefix = if xfrm_ns == ns::P { "p" } else { "a" };
                let x = doc.create_element(xfrm_ns, prefix, "xfrm", &[]);
                doc.insert_child(parent, 0, x);
                x
            }
        };
        let node = match doc.first_child_named(xfrm, ns::A, child) {
            Some(n) => n,
            None => {
                let n = doc.create_element(ns::A, "a", child, &[]);
                // schema order: a:off must precede a:ext
                if child == "off" {
                    doc.insert_child(xfrm, 0, n);
                } else {
                    doc.append_child(xfrm, n);
                }
                n
            }
        };
        doc.set_attr(node, attr, &value.to_string());
        Ok(())
    }

    fn text_frame_inner(&self, py: Python<'_>) -> PyResult<TextFrame> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let body = tx_body(doc, self.node)
            .ok_or_else(|| PyValueError::new_err("shape has no text frame"))?;
        drop(prs);
        Ok(TextFrame {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node: body,
        })
    }
}

#[pymethods]
impl Shape {
    #[getter]
    fn shape_id(&self, py: Python<'_>) -> PyResult<u32> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(c_nv_pr(doc, self.node)
            .and_then(|n| doc.attr(n, "id"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(0))
    }

    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<String> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(c_nv_pr(doc, self.node)
            .and_then(|n| doc.attr(n, "name"))
            .unwrap_or_default()
            .to_string())
    }

    #[getter]
    fn has_text_frame(&self, py: Python<'_>) -> PyResult<bool> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(tx_body(doc, self.node).is_some())
    }

    #[getter]
    fn text_frame(&self, py: Python<'_>) -> PyResult<TextFrame> {
        self.text_frame_inner(py)
    }

    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        self.text_frame_inner(py)?.get_text(py)
    }

    #[setter(text)]
    fn set_text(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.text_frame_inner(py)?.set_text(py, value)
    }

    #[getter]
    fn left(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        self.xfrm_val(py, "off", "x")
    }

    #[setter(left)]
    fn set_left(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        self.set_xfrm_val(py, "off", "x", value)
    }

    #[getter]
    fn top(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        self.xfrm_val(py, "off", "y")
    }

    #[setter(top)]
    fn set_top(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        self.set_xfrm_val(py, "off", "y", value)
    }

    #[getter]
    fn width(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        self.xfrm_val(py, "ext", "cx")
    }

    #[setter(width)]
    fn set_width(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        self.set_xfrm_val(py, "ext", "cx", value)
    }

    #[getter]
    fn height(&self, py: Python<'_>) -> PyResult<Option<i64>> {
        self.xfrm_val(py, "ext", "cy")
    }

    #[setter(height)]
    fn set_height(&self, py: Python<'_>, value: i64) -> PyResult<()> {
        self.set_xfrm_val(py, "ext", "cy", value)
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other.extract::<PyRef<'_, Shape>>().is_ok_and(|o| {
            self.prs.as_ptr() == o.prs.as_ptr() && self.part == o.part && self.node == o.node
        })
    }
}

// ---------------------------------------------------------------------------
// TextFrame
// ---------------------------------------------------------------------------

#[pyclass(name = "TextFrame", module = "pptx_rs._core")]
pub struct TextFrame {
    prs: Py<Presentation>,
    part: String,
    /// `p:txBody`
    node: NodeId,
}

fn paragraph_text(doc: &Document, p: NodeId) -> String {
    let mut out = String::new();
    for child in doc.child_elements(p) {
        let el = doc.get(child);
        if el.ns.as_deref() != Some(ns::A) {
            continue;
        }
        match el.local.as_str() {
            "r" | "fld" => {
                if let Some(t) = doc.first_child_named(child, ns::A, "t") {
                    out.push_str(&doc.text(t));
                }
            }
            "br" => out.push(VERTICAL_TAB),
            _ => {}
        }
    }
    out
}

/// Characters XML 1.0 cannot represent are stored as `_xHHHH_` escapes,
/// matching python-pptx (e.g. `\x1b` → `_x001B_`).
fn escape_ctrl_chars(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for c in text.chars() {
        match c {
            '\t' | '\n' | '\r' => out.push(c),
            c if (c as u32) < 0x20 || c as u32 == 0x7F => {
                out.push_str(&format!("_x{:04X}_", c as u32));
            }
            c => out.push(c),
        }
    }
    out
}

/// Replace a paragraph's content (runs/breaks) with `text`. A paragraph cannot
/// hold a newline, so both `\n` and `\v` become `a:br` soft line-breaks
/// (python-pptx semantics). `a:pPr` and `a:endParaRPr` children are preserved.
fn set_paragraph_text(doc: &mut Document, p: NodeId, text: &str) {
    for child in doc.child_elements(p) {
        let el = doc.get(child);
        if el.ns.as_deref() == Some(ns::A) && matches!(el.local.as_str(), "r" | "br" | "fld") {
            doc.remove_child(p, child);
        }
    }
    for (i, segment) in text.split(['\n', VERTICAL_TAB]).enumerate() {
        if i > 0 {
            let br = doc.create_element(ns::A, "a", "br", &[]);
            insert_before_end_para_r_pr(doc, p, br);
        }
        if !segment.is_empty() {
            let r = new_run(doc, &escape_ctrl_chars(segment));
            insert_before_end_para_r_pr(doc, p, r);
        }
    }
}

fn new_run(doc: &mut Document, text: &str) -> NodeId {
    let r = doc.create_element(ns::A, "a", "r", &[]);
    let t = doc.create_element(ns::A, "a", "t", &[]);
    doc.set_text(t, text);
    doc.append_child(r, t);
    r
}

/// `a:endParaRPr` must stay the last child of `a:p` per the schema.
fn insert_before_end_para_r_pr(doc: &mut Document, p: NodeId, node: NodeId) {
    match doc.first_child_named(p, ns::A, "endParaRPr") {
        Some(end) => {
            let index = doc
                .child_elements(p)
                .into_iter()
                .position(|c| c == end)
                .expect("endParaRPr is a child of p");
            doc.insert_child(p, index, node);
        }
        None => doc.append_child(p, node),
    }
}

impl TextFrame {
    fn get_text(&self, py: Python<'_>) -> PyResult<String> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let texts: Vec<String> = doc
            .children_named(self.node, ns::A, "p")
            .into_iter()
            .map(|p| paragraph_text(doc, p))
            .collect();
        Ok(texts.join("\n"))
    }

    fn set_text(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc_mut(&self.part)?;
        let paras = doc.children_named(self.node, ns::A, "p");
        // one paragraph per line: reuse the first, drop the rest, append new
        for &p in paras.iter().skip(1) {
            doc.remove_child(self.node, p);
        }
        let mut lines = value.split('\n');
        let first_line = lines.next().unwrap_or_default();
        let first_p = match paras.first() {
            Some(&p) => p,
            None => {
                let p = doc.create_element(ns::A, "a", "p", &[]);
                doc.append_child(self.node, p);
                p
            }
        };
        set_paragraph_text(doc, first_p, first_line);
        for line in lines {
            let p = doc.create_element(ns::A, "a", "p", &[]);
            set_paragraph_text(doc, p, line);
            doc.append_child(self.node, p);
        }
        Ok(())
    }
}

#[pymethods]
impl TextFrame {
    #[getter]
    fn paragraphs(&self, py: Python<'_>) -> PyResult<Vec<Paragraph>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let nodes = doc.children_named(self.node, ns::A, "p");
        drop(prs);
        Ok(nodes
            .into_iter()
            .map(|n| Paragraph {
                prs: self.prs.clone_ref(py),
                part: self.part.clone(),
                node: n,
            })
            .collect())
    }

    #[getter(text)]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        self.get_text(py)
    }

    #[setter(text)]
    fn text_setter(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        self.set_text(py, value)
    }

    fn add_paragraph(&self, py: Python<'_>) -> PyResult<Paragraph> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc_mut(&self.part)?;
        let p = doc.create_element(ns::A, "a", "p", &[]);
        doc.append_child(self.node, p);
        drop(prs);
        Ok(Paragraph {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node: p,
        })
    }
}

// ---------------------------------------------------------------------------
// Paragraph / Run
// ---------------------------------------------------------------------------

#[pyclass(name = "_Paragraph", module = "pptx_rs._core")]
pub struct Paragraph {
    prs: Py<Presentation>,
    part: String,
    /// `a:p`
    node: NodeId,
}

#[pymethods]
impl Paragraph {
    #[getter]
    fn runs(&self, py: Python<'_>) -> PyResult<Vec<Run>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let nodes = doc.children_named(self.node, ns::A, "r");
        drop(prs);
        Ok(nodes
            .into_iter()
            .map(|n| Run {
                prs: self.prs.clone_ref(py),
                part: self.part.clone(),
                node: n,
            })
            .collect())
    }

    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(paragraph_text(doc, self.node))
    }

    #[setter(text)]
    fn set_text(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc_mut(&self.part)?;
        set_paragraph_text(doc, self.node, value);
        Ok(())
    }

    fn add_run(&self, py: Python<'_>) -> PyResult<Run> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc_mut(&self.part)?;
        let r = new_run(doc, "");
        insert_before_end_para_r_pr(doc, self.node, r);
        drop(prs);
        Ok(Run {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node: r,
        })
    }
}

#[pyclass(name = "_Run", module = "pptx_rs._core")]
pub struct Run {
    prs: Py<Presentation>,
    part: String,
    /// `a:r`
    node: NodeId,
}

#[pymethods]
impl Run {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(doc
            .first_child_named(self.node, ns::A, "t")
            .map(|t| doc.text(t))
            .unwrap_or_default())
    }

    #[setter(text)]
    fn set_text(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc_mut(&self.part)?;
        let t = match doc.first_child_named(self.node, ns::A, "t") {
            Some(t) => t,
            None => {
                let t = doc.create_element(ns::A, "a", "t", &[]);
                doc.append_child(self.node, t);
                t
            }
        };
        doc.set_text(t, &escape_ctrl_chars(value));
        Ok(())
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other.extract::<PyRef<'_, Run>>().is_ok_and(|o| {
            self.prs.as_ptr() == o.prs.as_ptr() && self.part == o.part && self.node == o.node
        })
    }
}

// ---------------------------------------------------------------------------
// module
// ---------------------------------------------------------------------------

#[pymodule]
fn _core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add_class::<Presentation>()?;
    m.add_class::<Slides>()?;
    m.add_class::<Slide>()?;
    m.add_class::<SlideShapes>()?;
    m.add_class::<Shape>()?;
    m.add_class::<TextFrame>()?;
    m.add_class::<Paragraph>()?;
    m.add_class::<Run>()?;
    Ok(())
}
