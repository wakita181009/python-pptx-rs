//! Chart read API over a chart part (`c:chartSpace`).

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyTuple;

use pptx_core::dom::{Document, NodeId};
use pptx_core::ns;

use crate::{Presentation, TextFrame};

/// First `ns:local` descendant of `root` (excluding `root` itself), returning
/// as soon as it is found (depth-first, document order).
fn first_descendant(doc: &Document, root: NodeId, ns_uri: &str, local: &str) -> Option<NodeId> {
    for child in doc.child_elements(root) {
        if doc.is(child, ns_uri, local) {
            return Some(child);
        }
        if let Some(found) = first_descendant(doc, child, ns_uri, local) {
            return Some(found);
        }
    }
    None
}

/// `c:chart` child of the `c:chartSpace` document root.
fn chart_elm(doc: &Document) -> PyResult<NodeId> {
    doc.first_child_named(doc.root, ns::C, "chart")
        .ok_or_else(|| PyValueError::new_err("chart part has no c:chart element"))
}

fn plot_area(doc: &Document) -> PyResult<NodeId> {
    let chart = chart_elm(doc)?;
    doc.first_child_named(chart, ns::C, "plotArea")
        .ok_or_else(|| PyValueError::new_err("chart has no c:plotArea element"))
}

/// The `c:*Chart` plot elements of the plot area, in document order.
fn x_charts(doc: &Document) -> PyResult<Vec<NodeId>> {
    let area = plot_area(doc)?;
    Ok(doc
        .child_elements(area)
        .into_iter()
        .filter(|&n| {
            let el = doc.get(n);
            el.ns.as_deref() == Some(ns::C) && el.local.ends_with("Chart")
        })
        .collect())
}

/// Point values of a cache element (`c:strCache`, `c:numCache`, or
/// `c:multiLvlStrCache`) in leaf-category order, `None` where a `c:pt` is
/// absent for that index. Multi-level caches keep their leaf labels in the
/// first `c:lvl`. Reads every `c:pt` in a single pass (`c:ptCount` sets the
/// length; points may be sparse and out of order via their `idx`).
fn cache_texts(doc: &Document, cache: NodeId) -> Vec<Option<String>> {
    let container = doc.first_child_named(cache, ns::C, "lvl").unwrap_or(cache);
    let count: usize = doc
        .first_child_named(cache, ns::C, "ptCount")
        .and_then(|n| doc.attr(n, "val"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    let mut texts = vec![None; count];
    for pt in doc.children_named(container, ns::C, "pt") {
        let Some(idx) = doc.attr(pt, "idx").and_then(|v| v.parse::<usize>().ok()) else {
            continue;
        };
        if idx < count
            && let Some(v) = doc.first_child_named(pt, ns::C, "v")
        {
            texts[idx] = Some(doc.text(v));
        }
    }
    texts
}

#[pyclass(name = "Chart", module = "pptx_rs._core")]
pub struct Chart {
    pub(crate) prs: Py<Presentation>,
    /// Chart part name (e.g. `ppt/charts/chart1.xml`).
    pub(crate) part: String,
}

#[pymethods]
impl Chart {
    #[getter]
    fn has_title(&self, py: Python<'_>) -> PyResult<bool> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let chart = chart_elm(doc)?;
        Ok(doc.first_child_named(chart, ns::C, "title").is_some())
    }

    #[getter]
    fn chart_title(&self, py: Python<'_>) -> ChartTitle {
        ChartTitle {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
        }
    }

    #[getter]
    fn plots(&self, py: Python<'_>) -> PyResult<Vec<Plot>> {
        let mut prs = self.prs.borrow_mut(py);
        let nodes = x_charts(prs.doc(&self.part)?)?;
        drop(prs);
        Ok(nodes
            .into_iter()
            .map(|n| Plot {
                prs: self.prs.clone_ref(py),
                part: self.part.clone(),
                node: n,
            })
            .collect())
    }

    /// All series in the chart, in plot then document order (python-pptx
    /// `chart.series` semantics).
    #[getter]
    fn series(&self, py: Python<'_>) -> PyResult<Vec<Series>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let mut nodes = Vec::new();
        for x_chart in x_charts(doc)? {
            nodes.extend(doc.children_named(x_chart, ns::C, "ser"));
        }
        drop(prs);
        Ok(nodes
            .into_iter()
            .map(|n| Series {
                prs: self.prs.clone_ref(py),
                part: self.part.clone(),
                node: n,
            })
            .collect())
    }
}

#[pyclass(name = "ChartTitle", module = "pptx_rs._core")]
pub struct ChartTitle {
    prs: Py<Presentation>,
    part: String,
}

#[pymethods]
impl ChartTitle {
    /// The title's rich-text frame. When the `c:title` carries no `c:tx/c:rich`
    /// (e.g. a title bound to a cell via `c:strRef`), python-pptx returns an
    /// empty text frame; this read-only engine points the frame at the title
    /// element itself, which has no `a:p` children, so `.text` reads `""`
    /// without mutating the part. (python-pptx would materialize the `c:rich`.)
    #[getter]
    fn text_frame(&self, py: Python<'_>) -> PyResult<TextFrame> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let chart = chart_elm(doc)?;
        let title = doc
            .first_child_named(chart, ns::C, "title")
            .ok_or_else(|| PyValueError::new_err("chart has no c:title element"))?;
        let node = doc
            .first_child_named(title, ns::C, "tx")
            .and_then(|tx| doc.first_child_named(tx, ns::C, "rich"))
            .unwrap_or(title);
        drop(prs);
        Ok(TextFrame {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node,
        })
    }
}

#[pyclass(name = "_BasePlot", module = "pptx_rs._core")]
pub struct Plot {
    prs: Py<Presentation>,
    part: String,
    /// A `c:*Chart` plot element (e.g. `c:barChart`).
    node: NodeId,
}

#[pymethods]
impl Plot {
    /// Categories of the first series, one per leaf category; a category with
    /// no `c:pt` yields an empty label (python-pptx semantics).
    #[getter]
    fn categories(&self, py: Python<'_>) -> PyResult<Vec<Category>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let Some(cat) = doc
            .children_named(self.node, ns::C, "ser")
            .first()
            .and_then(|&ser| doc.first_child_named(ser, ns::C, "cat"))
        else {
            return Ok(Vec::new());
        };
        let Some(cache) = ["strCache", "numCache", "multiLvlStrCache"]
            .iter()
            .find_map(|local| first_descendant(doc, cat, ns::C, local))
        else {
            return Ok(Vec::new());
        };
        Ok(cache_texts(doc, cache)
            .into_iter()
            .map(|label| Category {
                label: label.unwrap_or_default(),
            })
            .collect())
    }
}

/// python-pptx's `Category` is a `str` subclass; only the `label` accessor is
/// implemented here.
#[pyclass(name = "Category", module = "pptx_rs._core")]
pub struct Category {
    #[pyo3(get)]
    label: String,
}

#[pymethods]
impl Category {
    fn __str__(&self) -> String {
        self.label.clone()
    }

    fn __repr__(&self) -> String {
        format!("{:?}", self.label)
    }
}

#[pyclass(name = "Series", module = "pptx_rs._core")]
pub struct Series {
    prs: Py<Presentation>,
    part: String,
    /// `c:ser`
    node: NodeId,
}

#[pymethods]
impl Series {
    #[getter]
    fn name(&self, py: Python<'_>) -> PyResult<String> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        Ok(doc
            .first_child_named(self.node, ns::C, "tx")
            .and_then(|tx| first_descendant(doc, tx, ns::C, "pt"))
            .and_then(|pt| doc.first_child_named(pt, ns::C, "v"))
            .map(|v| doc.text(v))
            .unwrap_or_default())
    }

    /// Series values in chart order; a data point with no cached value is
    /// None (python-pptx returns a tuple, matched here).
    #[getter]
    fn values<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyTuple>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let values: Vec<Option<f64>> = match doc.first_child_named(self.node, ns::C, "val") {
            Some(val) => {
                let cache = ["numCache", "numLit"]
                    .iter()
                    .find_map(|local| first_descendant(doc, val, ns::C, local));
                match cache {
                    Some(cache) => cache_texts(doc, cache)
                        .into_iter()
                        .map(|t| t.and_then(|t| t.parse().ok()))
                        .collect(),
                    None => Vec::new(),
                }
            }
            None => Vec::new(),
        };
        PyTuple::new(py, values)
    }
}
