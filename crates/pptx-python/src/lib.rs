//! PyO3 bindings exposing a python-pptx-compatible API.
//!
//! Every proxy object holds `Py<Presentation>` plus stable node ids into the
//! owning part's DOM, so one attribute access is a single Rust call instead of
//! python-pptx's descriptor/lxml call chain.

mod chart;
mod enums;
mod image;
mod table;

use std::collections::HashMap;
use std::path::PathBuf;

use pyo3::exceptions::{
    PyAttributeError, PyIndexError, PyKeyError, PyNotImplementedError, PyValueError,
};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyIterator, PyList};

use pptx_core::dom::{Document, NodeId};
use pptx_core::error::Error;
use pptx_core::ns;
use pptx_core::opc::Package;
use pyenum::PyModuleExt;

use crate::enums::MSO_SHAPE_TYPE;

pub(crate) fn to_py(e: Error) -> PyErr {
    match e {
        Error::Io(e) => e.into(),
        Error::PartNotFound(p) => PyKeyError::new_err(p),
        other => PyValueError::new_err(other.to_string()),
    }
}

const OFFICE_DOCUMENT_RELTYPE_SUFFIX: &str = "/officeDocument";
const SLIDE_MASTER_RELTYPE_SUFFIX: &str = "/slideMaster";
const SLIDE_LAYOUT_RELTYPE_SUFFIX: &str = "/slideLayout";
const NOTES_SLIDE_RELTYPE_SUFFIX: &str = "/notesSlide";
/// `a:graphicData/@uri` values distinguishing graphic-frame payloads
/// (the chart URI is the chart namespace itself, `ns::C`).
const GRAPHIC_DATA_URI_TABLE: &str = "http://schemas.openxmlformats.org/drawingml/2006/table";
const GRAPHIC_DATA_URI_OLE: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";
const SLIDE_RELTYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_RELTYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
/// Base XML for a new slide part (python-pptx's CT_Slide.new template).
const NEW_SLIDE_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    "\r\n",
    r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
    r#"<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>"#,
    r#"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"#,
    r#"</p:sld>"#,
);
/// Layout placeholders NOT cloned onto a new slide (python-pptx semantics).
const LATENT_PH_TYPES: [&str; 3] = ["dt", "ftr", "sldNum"];
/// Placeholder types that get an empty text body on the new slide.
const TEXT_PH_TYPES: [&str; 5] = ["title", "ctrTitle", "subTitle", "body", "obj"];
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

    pub(crate) fn doc(&mut self, part: &str) -> PyResult<&Document> {
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

    /// Part names referenced from `id_lst` children (`p:sldMasterId`,
    /// `p:sldLayoutId`, ...) of `part`, in document order.
    fn rel_id_list_parts(
        &mut self,
        part: &str,
        lst_local: &str,
        id_local: &str,
    ) -> PyResult<Vec<String>> {
        let rels = self.pkg.rels(part).map_err(to_py)?;
        let rel_targets: HashMap<String, String> = rels
            .into_iter()
            .map(|r| (r.id, self.pkg.resolve_target(part, &r.target)))
            .collect();
        let doc = self.pkg.doc(part).map_err(to_py)?;
        let Some(lst) = doc.first_child_named(doc.root, ns::P, lst_local) else {
            return Ok(Vec::new());
        };
        Ok(doc
            .children_named(lst, ns::P, id_local)
            .into_iter()
            .filter_map(|n| doc.attr(n, "r:id"))
            .filter_map(|rid| rel_targets.get(rid).cloned())
            .collect())
    }

    fn master_parts(&mut self) -> PyResult<Vec<String>> {
        let part = self.part_name.clone();
        self.rel_id_list_parts(&part, "sldMasterIdLst", "sldMasterId")
    }

    fn layout_parts(&mut self, master_part: &str) -> PyResult<Vec<String>> {
        self.rel_id_list_parts(master_part, "sldLayoutIdLst", "sldLayoutId")
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
    fn slide_masters(slf: Py<Self>) -> SlideMasters {
        SlideMasters { prs: slf }
    }

    /// Layouts of the first slide master (python-pptx convenience shortcut).
    #[getter]
    fn slide_layouts(slf: Py<Self>, py: Python<'_>) -> PyResult<SlideLayouts> {
        let master_part = {
            let mut prs = slf.borrow_mut(py);
            prs.master_parts()?
                .into_iter()
                .next()
                .ok_or_else(|| PyValueError::new_err("presentation has no slide master"))?
        };
        Ok(SlideLayouts {
            prs: slf,
            master_part,
        })
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

    fn add_slide(&self, py: Python<'_>, slide_layout: PyRef<'_, SlideLayout>) -> PyResult<Slide> {
        let layout_part = slide_layout.part.clone();
        let mut prs = self.prs.borrow_mut(py);

        // placeholder specs from the layout: (name, ph attrs), latent ones skipped
        let ph_specs: Vec<(String, Vec<(String, String)>)> = {
            let ldoc = prs.doc(&layout_part)?;
            let tree = sp_tree(ldoc)?;
            ldoc.children_named(tree, ns::P, "sp")
                .into_iter()
                .filter_map(|sp| {
                    let nv_sp_pr = ldoc.first_child_named(sp, ns::P, "nvSpPr")?;
                    let nv_pr = ldoc.first_child_named(nv_sp_pr, ns::P, "nvPr")?;
                    let ph = ldoc.first_child_named(nv_pr, ns::P, "ph")?;
                    let ph_type = ldoc.attr(ph, "type").unwrap_or("obj");
                    if LATENT_PH_TYPES.contains(&ph_type) {
                        return None;
                    }
                    let name = c_nv_pr(ldoc, sp)
                        .and_then(|n| ldoc.attr(n, "name"))
                        .unwrap_or_default()
                        .to_string();
                    Some((name, ldoc.get(ph).attrs.clone()))
                })
                .collect()
        };

        let mut sdoc = Document::parse(NEW_SLIDE_XML.as_bytes()).map_err(to_py)?;
        let tree = sp_tree(&sdoc)?;
        let mut next_id = 2u32;
        for (name, ph_attrs) in &ph_specs {
            let sp = new_placeholder_sp(&mut sdoc, next_id, name, ph_attrs);
            sdoc.append_child(tree, sp);
            next_id += 1;
        }

        let part = next_slide_part_name(prs.pkg.part_names());
        prs.pkg
            .add_xml_part(&part, SLIDE_CONTENT_TYPE, sdoc)
            .map_err(to_py)?;
        prs.next_shape_ids.insert(part.clone(), next_id);

        let layout_target = pptx_core::opc::relative_target(&part, &layout_part);
        prs.pkg
            .add_relationship(&part, SLIDE_LAYOUT_RELTYPE, &layout_target)
            .map_err(to_py)?;

        let pres_part = prs.part_name.clone();
        let slide_target = pptx_core::opc::relative_target(&pres_part, &part);
        let rid = prs
            .pkg
            .add_relationship(&pres_part, SLIDE_RELTYPE, &slide_target)
            .map_err(to_py)?;

        // python-pptx assigns slide ids starting at 256
        let slide_id = prs
            .slide_entries
            .iter()
            .map(|(_, id)| *id)
            .max()
            .unwrap_or(255)
            .max(255)
            + 1;
        let pdoc = prs.doc_mut(&pres_part)?;
        let sld_id_lst = get_or_add_sld_id_lst(pdoc)?;
        let sld_id = pdoc.create_element(
            ns::P,
            "p",
            "sldId",
            &[("id", slide_id.to_string().as_str()), ("r:id", &rid)],
        );
        pdoc.append_child(sld_id_lst, sld_id);
        prs.slide_entries.push((part.clone(), slide_id));
        drop(prs);

        Ok(Slide {
            prs: self.prs.clone_ref(py),
            part,
            slide_id,
        })
    }
}

/// Next free `ppt/slides/slideN.xml` part name.
fn next_slide_part_name(part_names: &[String]) -> String {
    let max_n = part_names
        .iter()
        .filter_map(|n| {
            n.strip_prefix("ppt/slides/slide")
                .and_then(|s| s.strip_suffix(".xml"))
                .and_then(|s| s.parse::<u32>().ok())
        })
        .max()
        .unwrap_or(0);
    format!("ppt/slides/slide{}.xml", max_n + 1)
}

/// Get `p:sldIdLst`, creating it right after `p:sldMasterIdLst` when missing
/// (the schema orders it before `p:sldSz`).
fn get_or_add_sld_id_lst(doc: &mut Document) -> PyResult<NodeId> {
    if let Some(lst) = doc.first_child_named(doc.root, ns::P, "sldIdLst") {
        return Ok(lst);
    }
    let root = doc.root;
    let lst = doc.create_element(ns::P, "p", "sldIdLst", &[]);
    let index = doc
        .child_elements(root)
        .into_iter()
        .position(|c| doc.is(c, ns::P, "sldMasterIdLst"))
        .map(|i| i + 1)
        .unwrap_or(0);
    doc.insert_child(root, index, lst);
    Ok(lst)
}

/// Build a fresh placeholder `p:sp` (python-pptx's `new_placeholder_sp`),
/// copying the layout placeholder's `p:ph` attributes verbatim.
fn new_placeholder_sp(
    doc: &mut Document,
    shape_id: u32,
    name: &str,
    ph_attrs: &[(String, String)],
) -> NodeId {
    let id_s = shape_id.to_string();
    let sp = doc.create_element(ns::P, "p", "sp", &[]);

    let nv_sp_pr = doc.create_element(ns::P, "p", "nvSpPr", &[]);
    let c_nv_pr = doc.create_element(
        ns::P,
        "p",
        "cNvPr",
        &[("id", id_s.as_str()), ("name", name)],
    );
    let c_nv_sp_pr = doc.create_element(ns::P, "p", "cNvSpPr", &[]);
    let sp_locks = doc.create_element(ns::A, "a", "spLocks", &[("noGrp", "1")]);
    doc.append_child(c_nv_sp_pr, sp_locks);
    let nv_pr = doc.create_element(ns::P, "p", "nvPr", &[]);
    let ph = doc.create_element(ns::P, "p", "ph", &[]);
    for (k, v) in ph_attrs {
        doc.set_attr(ph, k, v);
    }
    doc.append_child(nv_pr, ph);
    doc.append_child(nv_sp_pr, c_nv_pr);
    doc.append_child(nv_sp_pr, c_nv_sp_pr);
    doc.append_child(nv_sp_pr, nv_pr);
    doc.append_child(sp, nv_sp_pr);

    let sp_pr = doc.create_element(ns::P, "p", "spPr", &[]);
    doc.append_child(sp, sp_pr);

    let ph_type = ph_attrs
        .iter()
        .find(|(k, _)| k == "type")
        .map(|(_, v)| v.as_str())
        .unwrap_or("obj");
    if TEXT_PH_TYPES.contains(&ph_type) {
        let tx_body = doc.create_element(ns::P, "p", "txBody", &[]);
        let body_pr = doc.create_element(ns::A, "a", "bodyPr", &[]);
        let lst_style = doc.create_element(ns::A, "a", "lstStyle", &[]);
        let p = doc.create_element(ns::A, "a", "p", &[]);
        doc.append_child(tx_body, body_pr);
        doc.append_child(tx_body, lst_style);
        doc.append_child(tx_body, p);
        doc.append_child(sp, tx_body);
    }

    sp
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
            container: None,
        }
    }

    #[getter]
    fn has_notes_slide(&self, py: Python<'_>) -> PyResult<bool> {
        Ok(self.notes_part(py)?.is_some())
    }

    #[getter]
    fn notes_slide(&self, py: Python<'_>) -> PyResult<NotesSlide> {
        // python-pptx creates a notes slide on first access; that is a write
        // path this engine does not implement yet.
        let part = self.notes_part(py)?.ok_or_else(|| {
            PyValueError::new_err("slide has no notes slide (creating one is not implemented)")
        })?;
        Ok(NotesSlide {
            prs: self.prs.clone_ref(py),
            part,
        })
    }

    #[getter]
    fn slide_layout(&self, py: Python<'_>) -> PyResult<SlideLayout> {
        let mut prs = self.prs.borrow_mut(py);
        let rels = prs.pkg.rels(&self.part).map_err(to_py)?;
        let layout_part = rels
            .iter()
            .find(|r| r.reltype.ends_with(SLIDE_LAYOUT_RELTYPE_SUFFIX))
            .map(|r| prs.pkg.resolve_target(&self.part, &r.target))
            .ok_or_else(|| PyValueError::new_err("slide has no slide-layout relationship"))?;
        drop(prs);
        Ok(SlideLayout {
            prs: self.prs.clone_ref(py),
            part: layout_part,
        })
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .extract::<PyRef<'_, Slide>>()
            .is_ok_and(|o| self.prs.as_ptr() == o.prs.as_ptr() && self.part == o.part)
    }
}

impl Slide {
    fn notes_part(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let mut prs = self.prs.borrow_mut(py);
        let rels = prs.pkg.rels(&self.part).map_err(to_py)?;
        Ok(rels
            .iter()
            .find(|r| r.reltype.ends_with(NOTES_SLIDE_RELTYPE_SUFFIX))
            .map(|r| prs.pkg.resolve_target(&self.part, &r.target)))
    }
}

// ---------------------------------------------------------------------------
// NotesSlide
// ---------------------------------------------------------------------------

#[pyclass(module = "pptx_rs._core")]
pub struct NotesSlide {
    prs: Py<Presentation>,
    part: String,
}

#[pymethods]
impl NotesSlide {
    /// Text frame of the notes (body) placeholder, or None when the notes
    /// slide has no body placeholder (python-pptx semantics).
    #[getter]
    fn notes_text_frame(&self, py: Python<'_>) -> PyResult<Option<TextFrame>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let tree = sp_tree(doc)?;
        let body = shape_children(doc, tree).into_iter().find(|&n| {
            ph_elm(doc, n).is_some_and(|ph| doc.attr(ph, "type").unwrap_or("obj") == "body")
        });
        let tx = body.and_then(|n| tx_body(doc, n));
        drop(prs);
        Ok(tx.map(|node| TextFrame {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node,
        }))
    }
}

// ---------------------------------------------------------------------------
// SlideMasters / SlideMaster / SlideLayouts / SlideLayout
// ---------------------------------------------------------------------------

#[pyclass(module = "pptx_rs._core")]
pub struct SlideMasters {
    prs: Py<Presentation>,
}

#[pymethods]
impl SlideMasters {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.prs.borrow_mut(py).master_parts()?.len())
    }

    fn __getitem__(&self, py: Python<'_>, idx: isize) -> PyResult<SlideMaster> {
        let parts = self.prs.borrow_mut(py).master_parts()?;
        let i = normalize_index(idx, parts.len(), "slide master")?;
        Ok(SlideMaster {
            prs: self.prs.clone_ref(py),
            part: parts[i].clone(),
        })
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyIterator>> {
        let masters: Vec<SlideMaster> = self
            .prs
            .borrow_mut(py)
            .master_parts()?
            .into_iter()
            .map(|part| SlideMaster {
                prs: self.prs.clone_ref(py),
                part,
            })
            .collect();
        Ok(PyList::new(py, masters)?.try_iter()?.unbind())
    }
}

#[pyclass(module = "pptx_rs._core")]
pub struct SlideMaster {
    prs: Py<Presentation>,
    part: String,
}

#[pymethods]
impl SlideMaster {
    #[getter]
    fn slide_layouts(&self, py: Python<'_>) -> SlideLayouts {
        SlideLayouts {
            prs: self.prs.clone_ref(py),
            master_part: self.part.clone(),
        }
    }
}

#[pyclass(module = "pptx_rs._core")]
pub struct SlideLayouts {
    prs: Py<Presentation>,
    master_part: String,
}

impl SlideLayouts {
    fn parts(&self, py: Python<'_>) -> PyResult<Vec<String>> {
        self.prs.borrow_mut(py).layout_parts(&self.master_part)
    }

    fn layout(&self, py: Python<'_>, part: String) -> SlideLayout {
        SlideLayout {
            prs: self.prs.clone_ref(py),
            part,
        }
    }
}

#[pymethods]
impl SlideLayouts {
    fn __len__(&self, py: Python<'_>) -> PyResult<usize> {
        Ok(self.parts(py)?.len())
    }

    fn __getitem__(&self, py: Python<'_>, idx: isize) -> PyResult<SlideLayout> {
        let parts = self.parts(py)?;
        let i = normalize_index(idx, parts.len(), "slide layout")?;
        Ok(self.layout(py, parts[i].clone()))
    }

    fn __iter__(&self, py: Python<'_>) -> PyResult<Py<PyIterator>> {
        let layouts: Vec<SlideLayout> = self
            .parts(py)?
            .into_iter()
            .map(|part| self.layout(py, part))
            .collect();
        Ok(PyList::new(py, layouts)?.try_iter()?.unbind())
    }

    /// The layout named `name`, or None (python-pptx API).
    #[pyo3(signature = (name, default=None))]
    fn get_by_name(
        &self,
        py: Python<'_>,
        name: &str,
        default: Option<Py<PyAny>>,
    ) -> PyResult<Option<Py<PyAny>>> {
        for part in self.parts(py)? {
            let layout = self.layout(py, part);
            if layout.name(py)? == name {
                return Ok(Some(Py::new(py, layout)?.into_any()));
            }
        }
        Ok(default)
    }

    /// Index of `slide_layout` in this collection (python-pptx API).
    fn index(&self, py: Python<'_>, slide_layout: PyRef<'_, SlideLayout>) -> PyResult<usize> {
        if slide_layout.prs.as_ptr() != self.prs.as_ptr() {
            return Err(PyValueError::new_err(
                "layout not in this SlideLayouts collection",
            ));
        }
        self.parts(py)?
            .iter()
            .position(|p| *p == slide_layout.part)
            .ok_or_else(|| PyValueError::new_err("layout not in this SlideLayouts collection"))
    }
}

#[pyclass(module = "pptx_rs._core")]
pub struct SlideLayout {
    prs: Py<Presentation>,
    part: String,
}

#[pymethods]
impl SlideLayout {
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
    fn slide_master(&self, py: Python<'_>) -> PyResult<SlideMaster> {
        let mut prs = self.prs.borrow_mut(py);
        let rels = prs.pkg.rels(&self.part).map_err(to_py)?;
        let master_part = rels
            .iter()
            .find(|r| r.reltype.ends_with(SLIDE_MASTER_RELTYPE_SUFFIX))
            .map(|r| prs.pkg.resolve_target(&self.part, &r.target))
            .ok_or_else(|| PyValueError::new_err("layout has no slide-master relationship"))?;
        drop(prs);
        Ok(SlideMaster {
            prs: self.prs.clone_ref(py),
            part: master_part,
        })
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .extract::<PyRef<'_, SlideLayout>>()
            .is_ok_and(|o| self.prs.as_ptr() == o.prs.as_ptr() && self.part == o.part)
    }
}

fn normalize_index(idx: isize, len: usize, what: &str) -> PyResult<usize> {
    let len = len as isize;
    let i = if idx < 0 { idx + len } else { idx };
    if i < 0 || i >= len {
        return Err(PyIndexError::new_err(format!("{what} index out of range")));
    }
    Ok(i as usize)
}

// ---------------------------------------------------------------------------
// SlideShapes
// ---------------------------------------------------------------------------

#[pyclass(module = "pptx_rs._core")]
pub struct SlideShapes {
    prs: Py<Presentation>,
    part: String,
    /// `p:grpSp` when this is a group's member collection; the part's
    /// `p:spTree` otherwise.
    container: Option<NodeId>,
}

fn sp_tree(doc: &Document) -> PyResult<NodeId> {
    doc.first_child_named(doc.root, ns::P, "cSld")
        .and_then(|c| doc.first_child_named(c, ns::P, "spTree"))
        .ok_or_else(|| PyValueError::new_err("slide has no p:spTree"))
}

fn shape_children(doc: &Document, parent: NodeId) -> Vec<NodeId> {
    doc.child_elements(parent)
        .into_iter()
        .filter(|&n| {
            let el = doc.get(n);
            el.ns.as_deref() == Some(ns::P) && SHAPE_TAGS.contains(&el.local.as_str())
        })
        .collect()
}

impl SlideShapes {
    fn nodes(&self, py: Python<'_>) -> PyResult<Vec<NodeId>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let parent = match self.container {
            Some(n) => n,
            None => sp_tree(doc)?,
        };
        Ok(shape_children(doc, parent))
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
        let tree = match self.container {
            Some(n) => n,
            None => sp_tree(doc)?,
        };
        let sp = new_textbox_sp(doc, shape_id, &name, left, top, width, height);
        doc.append_child(tree, sp);
        drop(prs);
        Ok(self.shape(py, sp))
    }

    /// The title placeholder shape (`p:ph` idx 0), or None.
    #[getter]
    fn title(&self, py: Python<'_>) -> PyResult<Option<Shape>> {
        let node = {
            let nodes = self.nodes(py)?;
            let mut prs = self.prs.borrow_mut(py);
            let doc = prs.doc(&self.part)?;
            nodes.into_iter().find(|&n| {
                ph_elm(doc, n).is_some_and(|ph| doc.attr(ph, "idx").unwrap_or("0") == "0")
            })
        };
        Ok(node.map(|n| self.shape(py, n)))
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

/// The `p:nv*Pr` child of a shape element (`p:nvSpPr`, `p:nvPicPr`, ...).
fn nv_xx_pr(doc: &Document, shape: NodeId) -> Option<NodeId> {
    doc.child_elements(shape).into_iter().find(|&child| {
        let el = doc.get(child);
        el.ns.as_deref() == Some(ns::P) && el.local.starts_with("nv") && el.local.ends_with("Pr")
    })
}

/// The `p:ph` element of a placeholder shape, or None for non-placeholders.
fn ph_elm(doc: &Document, shape: NodeId) -> Option<NodeId> {
    let nv = nv_xx_pr(doc, shape)?;
    let nv_pr = doc.first_child_named(nv, ns::P, "nvPr")?;
    doc.first_child_named(nv_pr, ns::P, "ph")
}

/// `a:graphicData` of a `p:graphicFrame`.
fn graphic_data(doc: &Document, frame: NodeId) -> Option<NodeId> {
    let graphic = doc.first_child_named(frame, ns::A, "graphic")?;
    doc.first_child_named(graphic, ns::A, "graphicData")
}

/// `a:graphicData/@uri` of a `p:graphicFrame`.
fn graphic_data_uri(doc: &Document, frame: NodeId) -> Option<String> {
    graphic_data(doc, frame).and_then(|gd| doc.attr(gd, "uri").map(str::to_string))
}

/// Master placeholder type a layout placeholder of `ph_type` inherits from
/// (python-pptx `LayoutPlaceholder._base_placeholder` mapping, in XML terms).
fn master_ph_base_type(ph_type: &str) -> Option<&'static str> {
    Some(match ph_type {
        "title" | "ctrTitle" => "title",
        "body" | "subTitle" | "obj" | "chart" | "tbl" | "clipArt" | "dgm" | "media" | "pic" => {
            "body"
        }
        "dt" => "dt",
        "ftr" => "ftr",
        "sldNum" => "sldNum",
        _ => return None,
    })
}

/// Directly-applied `a:xfrm` value of a shape, ignoring inheritance.
fn direct_xfrm_val(doc: &Document, node: NodeId, child: &str, attr: &str) -> Option<i64> {
    xfrm_parent(doc, node)
        .and_then(|(parent, xfrm_ns)| doc.first_child_named(parent, xfrm_ns, "xfrm"))
        .and_then(|xfrm| doc.first_child_named(xfrm, ns::A, child))
        .and_then(|n| doc.attr(n, attr))
        .and_then(|v| v.parse().ok())
}

/// Which parent a placeholder inherits geometry from. The level is tracked
/// explicitly rather than sniffed from relationships because a `slideMaster`
/// part *also* carries `/slideLayout` rels (to its own layouts); keying the
/// direction off those would make a master wrongly inherit from a layout.
#[derive(Clone, Copy)]
enum PhLevel {
    /// A slide placeholder: inherits from the layout placeholder with the
    /// same `idx`.
    Slide,
    /// A layout placeholder: inherits from the master placeholder with the
    /// mapped `type`. The master is the inheritance root (inherits nothing).
    Layout,
}

/// Effective `a:xfrm` value of a slide shape: its directly-applied value, or
/// for a placeholder the value inherited along slide → layout → master
/// (python-pptx semantics).
fn effective_xfrm_val(
    prs: &mut Presentation,
    part: &str,
    node: NodeId,
    child: &str,
    attr: &str,
) -> PyResult<Option<i64>> {
    inherit_xfrm_val(prs, part, node, child, attr, PhLevel::Slide)
}

fn inherit_xfrm_val(
    prs: &mut Presentation,
    part: &str,
    node: NodeId,
    child: &str,
    attr: &str,
    level: PhLevel,
) -> PyResult<Option<i64>> {
    let (idx, ph_type) = {
        let doc = prs.doc(part)?;
        if let Some(v) = direct_xfrm_val(doc, node, child, attr) {
            return Ok(Some(v));
        }
        match ph_elm(doc, node) {
            Some(ph) => (
                doc.attr(ph, "idx").unwrap_or("0").to_string(),
                doc.attr(ph, "type").unwrap_or("obj").to_string(),
            ),
            None => return Ok(None),
        }
    };
    let (rel_suffix, match_by_idx) = match level {
        PhLevel::Slide => (SLIDE_LAYOUT_RELTYPE_SUFFIX, true),
        PhLevel::Layout => (SLIDE_MASTER_RELTYPE_SUFFIX, false),
    };
    let rels = prs.pkg.rels(part).map_err(to_py)?;
    let Some(r) = rels.iter().find(|r| r.reltype.ends_with(rel_suffix)) else {
        return Ok(None);
    };
    let parent_part = prs.pkg.resolve_target(part, &r.target);
    let base_type = master_ph_base_type(&ph_type);
    let parent_node = {
        let doc = prs.doc(&parent_part)?;
        let tree = sp_tree(doc)?;
        shape_children(doc, tree).into_iter().find(|&n| {
            let Some(p) = ph_elm(doc, n) else {
                return false;
            };
            if match_by_idx {
                doc.attr(p, "idx").unwrap_or("0") == idx
            } else {
                base_type.is_some_and(|bt| doc.attr(p, "type").unwrap_or("obj") == bt)
            }
        })
    };
    match (parent_node, level) {
        (Some(n), PhLevel::Slide) => {
            inherit_xfrm_val(prs, &parent_part, n, child, attr, PhLevel::Layout)
        }
        // The master is the inheritance root: take its direct value only.
        (Some(n), PhLevel::Layout) => Ok(direct_xfrm_val(prs.doc(&parent_part)?, n, child, attr)),
        (None, _) => Ok(None),
    }
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
        effective_xfrm_val(&mut prs, &self.part, self.node, child, attr)
    }

    /// Whether this shape's element is `p:<local>`, under a single borrow and
    /// with no allocation.
    fn is_local(&self, py: Python<'_>, local: &str) -> PyResult<bool> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(doc.get(self.node).local == local)
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

    /// python-pptx `MSO_SHAPE_TYPE` member (None for e.g. SmartArt frames).
    #[getter]
    fn shape_type(&self, py: Python<'_>) -> PyResult<Option<MSO_SHAPE_TYPE>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        match doc.get(self.node).local.as_str() {
            "sp" => {
                if ph_elm(doc, self.node).is_some() {
                    return Ok(Some(MSO_SHAPE_TYPE::PLACEHOLDER));
                }
                let sp_pr = doc.first_child_named(self.node, ns::P, "spPr");
                if sp_pr.is_some_and(|s| doc.first_child_named(s, ns::A, "custGeom").is_some()) {
                    return Ok(Some(MSO_SHAPE_TYPE::FREEFORM));
                }
                let is_textbox = nv_xx_pr(doc, self.node)
                    .and_then(|nv| doc.first_child_named(nv, ns::P, "cNvSpPr"))
                    .and_then(|c| doc.attr(c, "txBox"))
                    .is_some_and(|v| v == "1" || v == "true");
                let has_prst_geom =
                    sp_pr.is_some_and(|s| doc.first_child_named(s, ns::A, "prstGeom").is_some());
                if has_prst_geom && !is_textbox {
                    return Ok(Some(MSO_SHAPE_TYPE::AUTO_SHAPE));
                }
                if is_textbox {
                    return Ok(Some(MSO_SHAPE_TYPE::TEXT_BOX));
                }
                Err(PyNotImplementedError::new_err(
                    "Shape instance of unrecognized shape type",
                ))
            }
            "pic" => {
                if ph_elm(doc, self.node).is_some() {
                    return Ok(Some(MSO_SHAPE_TYPE::PLACEHOLDER));
                }
                let is_video = nv_xx_pr(doc, self.node)
                    .and_then(|nv| doc.first_child_named(nv, ns::P, "nvPr"))
                    .is_some_and(|nv_pr| {
                        doc.first_child_named(nv_pr, ns::A, "videoFile").is_some()
                    });
                Ok(Some(if is_video {
                    MSO_SHAPE_TYPE::MEDIA
                } else {
                    MSO_SHAPE_TYPE::PICTURE
                }))
            }
            "graphicFrame" => Ok(match graphic_data_uri(doc, self.node).as_deref() {
                Some(ns::C) => Some(MSO_SHAPE_TYPE::CHART),
                Some(GRAPHIC_DATA_URI_TABLE) => Some(MSO_SHAPE_TYPE::TABLE),
                Some(GRAPHIC_DATA_URI_OLE) => Some(MSO_SHAPE_TYPE::EMBEDDED_OLE_OBJECT),
                _ => None,
            }),
            "grpSp" => Ok(Some(MSO_SHAPE_TYPE::GROUP)),
            "cxnSp" => Ok(Some(MSO_SHAPE_TYPE::LINE)),
            other => Err(PyNotImplementedError::new_err(format!(
                "shape_type not implemented for p:{other}"
            ))),
        }
    }

    #[getter]
    fn has_chart(&self, py: Python<'_>) -> PyResult<bool> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(doc.get(self.node).local == "graphicFrame"
            && graphic_data_uri(doc, self.node).as_deref() == Some(ns::C))
    }

    #[getter]
    fn chart(&self, py: Python<'_>) -> PyResult<chart::Chart> {
        if !self.is_local(py, "graphicFrame")? {
            return Err(PyAttributeError::new_err("shape has no chart attribute"));
        }
        let mut prs = self.prs.borrow_mut(py);
        let rid = {
            let doc = prs.doc(&self.part)?;
            graphic_data(doc, self.node)
                .and_then(|gd| doc.first_child_named(gd, ns::C, "chart"))
                .and_then(|c| doc.attr(c, "r:id"))
                .map(str::to_string)
                .ok_or_else(|| PyValueError::new_err("shape does not contain a chart"))?
        };
        let rels = prs.pkg.rels(&self.part).map_err(to_py)?;
        let chart_part = rels
            .iter()
            .find(|r| r.id == rid)
            .map(|r| prs.pkg.resolve_target(&self.part, &r.target))
            .ok_or_else(|| PyKeyError::new_err(rid))?;
        drop(prs);
        Ok(chart::Chart {
            prs: self.prs.clone_ref(py),
            part: chart_part,
        })
    }

    #[getter]
    fn table(&self, py: Python<'_>) -> PyResult<table::Table> {
        if !self.is_local(py, "graphicFrame")? {
            return Err(PyAttributeError::new_err("shape has no table attribute"));
        }
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let tbl = graphic_data(doc, self.node)
            .and_then(|gd| doc.first_child_named(gd, ns::A, "tbl"))
            .ok_or_else(|| PyValueError::new_err("shape does not contain a table"))?;
        drop(prs);
        Ok(table::Table {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node: tbl,
        })
    }

    /// The embedded image of a `p:pic` shape. Raises AttributeError on other
    /// shape kinds so `hasattr(shape, "image")` works as in python-pptx,
    /// where only picture classes define this property.
    #[getter]
    fn image(&self, py: Python<'_>) -> PyResult<image::Image> {
        if !self.is_local(py, "pic")? {
            return Err(PyAttributeError::new_err("shape has no image attribute"));
        }
        let mut prs = self.prs.borrow_mut(py);
        let rid = {
            let doc = prs.doc(&self.part)?;
            doc.first_child_named(self.node, ns::P, "blipFill")
                .and_then(|bf| doc.first_child_named(bf, ns::A, "blip"))
                .and_then(|b| doc.attr(b, "r:embed"))
                .map(str::to_string)
                .ok_or_else(|| PyValueError::new_err("shape has no embedded image"))?
        };
        let rels = prs.pkg.rels(&self.part).map_err(to_py)?;
        let image_part = rels
            .iter()
            .find(|r| r.id == rid)
            .map(|r| prs.pkg.resolve_target(&self.part, &r.target))
            .ok_or_else(|| PyKeyError::new_err(rid))?;
        let blob = prs
            .pkg
            .raw(&image_part)
            .ok_or_else(|| PyKeyError::new_err(image_part.clone()))?
            .to_vec();
        let ext = image_part
            .rsplit_once('.')
            .map(|(_, e)| e.to_ascii_lowercase())
            .unwrap_or_default();
        Ok(image::Image { blob, ext })
    }

    /// Member shapes of a `p:grpSp` group shape.
    #[getter]
    fn shapes(&self, py: Python<'_>) -> PyResult<SlideShapes> {
        if !self.is_local(py, "grpSp")? {
            return Err(PyAttributeError::new_err("shape has no shapes attribute"));
        }
        Ok(SlideShapes {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            container: Some(self.node),
        })
    }

    /// Read-only escape hatch mirroring python-pptx's `shape._element` oxml
    /// access (e.g. `_element._nvXxPr.cNvPr.attrib`).
    #[getter]
    fn _element(&self, py: Python<'_>) -> XmlElement {
        XmlElement {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node: self.node,
        }
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
// XmlElement (read-only oxml escape hatch)
// ---------------------------------------------------------------------------

#[pyclass(name = "_XmlElement", module = "pptx_rs._core")]
pub struct XmlElement {
    prs: Py<Presentation>,
    part: String,
    node: NodeId,
}

#[pymethods]
impl XmlElement {
    /// Child element lookup by local name, mirroring python-pptx oxml
    /// attribute access; `_nvXxPr` resolves the `p:nv*Pr` child whatever the
    /// shape kind, as in python-pptx.
    fn __getattr__(&self, py: Python<'_>, name: &str) -> PyResult<XmlElement> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let found = if name == "_nvXxPr" {
            nv_xx_pr(doc, self.node)
        } else {
            doc.child_elements(self.node)
                .into_iter()
                .find(|&c| doc.get(c).local == name)
        };
        let node = found.ok_or_else(|| {
            PyAttributeError::new_err(format!("element has no child element {name}"))
        })?;
        drop(prs);
        Ok(XmlElement {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node,
        })
    }

    /// Attribute map, as lxml's `.attrib` (namespace declarations excluded).
    #[getter]
    fn attrib(&self, py: Python<'_>) -> PyResult<HashMap<String, String>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(doc
            .get(self.node)
            .attrs
            .iter()
            .filter(|(k, _)| k != "xmlns" && !k.starts_with("xmlns:"))
            .cloned()
            .collect())
    }
}

// ---------------------------------------------------------------------------
// TextFrame
// ---------------------------------------------------------------------------

#[pyclass(name = "TextFrame", module = "pptx_rs._core")]
pub struct TextFrame {
    pub(crate) prs: Py<Presentation>,
    pub(crate) part: String,
    /// `p:txBody` (or any element whose `a:p` children are the paragraphs,
    /// e.g. a chart title's `c:rich`)
    pub(crate) node: NodeId,
}

pub(crate) fn paragraph_text(doc: &Document, p: NodeId) -> String {
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
    m.add_class::<SlideMasters>()?;
    m.add_class::<SlideMaster>()?;
    m.add_class::<SlideLayouts>()?;
    m.add_class::<SlideLayout>()?;
    m.add_class::<SlideShapes>()?;
    m.add_class::<Shape>()?;
    m.add_class::<TextFrame>()?;
    m.add_class::<Paragraph>()?;
    m.add_class::<Run>()?;
    m.add_class::<NotesSlide>()?;
    m.add_class::<XmlElement>()?;
    m.add_class::<table::Table>()?;
    m.add_class::<table::Row>()?;
    m.add_class::<table::Cell>()?;
    m.add_class::<chart::Chart>()?;
    m.add_class::<chart::ChartTitle>()?;
    m.add_class::<chart::Plot>()?;
    m.add_class::<chart::Category>()?;
    m.add_class::<chart::Series>()?;
    m.add_class::<image::Image>()?;
    m.add_enum::<MSO_SHAPE_TYPE>()?;
    Ok(())
}
