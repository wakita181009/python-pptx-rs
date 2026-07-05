//! Table read API over `a:tbl`: `Table` → `_Row` → `_Cell`.

use pyo3::prelude::*;

use pptx_core::dom::NodeId;
use pptx_core::ns;

use crate::{Presentation, paragraph_text};

#[pyclass(name = "Table", module = "pptx_rs._core")]
pub struct Table {
    pub(crate) prs: Py<Presentation>,
    pub(crate) part: String,
    /// `a:tbl`
    pub(crate) node: NodeId,
}

#[pymethods]
impl Table {
    #[getter]
    fn rows(&self, py: Python<'_>) -> PyResult<Vec<Row>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let nodes = doc.children_named(self.node, ns::A, "tr");
        drop(prs);
        Ok(nodes
            .into_iter()
            .map(|n| Row {
                prs: self.prs.clone_ref(py),
                part: self.part.clone(),
                node: n,
            })
            .collect())
    }
}

#[pyclass(name = "_Row", module = "pptx_rs._core")]
pub struct Row {
    prs: Py<Presentation>,
    part: String,
    /// `a:tr`
    node: NodeId,
}

#[pymethods]
impl Row {
    #[getter]
    fn cells(&self, py: Python<'_>) -> PyResult<Vec<Cell>> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let nodes = doc.children_named(self.node, ns::A, "tc");
        drop(prs);
        Ok(nodes
            .into_iter()
            .map(|n| Cell {
                prs: self.prs.clone_ref(py),
                part: self.part.clone(),
                node: n,
            })
            .collect())
    }
}

#[pyclass(name = "_Cell", module = "pptx_rs._core")]
pub struct Cell {
    prs: Py<Presentation>,
    part: String,
    /// `a:tc`
    node: NodeId,
}

#[pymethods]
impl Cell {
    #[getter]
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let Some(tx_body) = doc.first_child_named(self.node, ns::A, "txBody") else {
            return Ok(String::new());
        };
        let texts: Vec<String> = doc
            .children_named(tx_body, ns::A, "p")
            .into_iter()
            .map(|p| paragraph_text(doc, p))
            .collect();
        Ok(texts.join("\n"))
    }
}
