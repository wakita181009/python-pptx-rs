//! Table read API over `a:tbl`: `Table` → `_Row` → `_Cell`.

use pyo3::prelude::*;

use pptx_core::dom::NodeId;
use pptx_core::ns;

use crate::{Presentation, paragraph_text, set_tx_body_text};

#[pyclass(name = "Table", module = "pptx_rs._core")]
pub struct Table {
    pub(crate) prs: Py<Presentation>,
    pub(crate) part: String,
    /// `a:tbl`
    pub(crate) node: NodeId,
}

#[pymethods]
impl Table {
    /// python-pptx `Table.cell(row_idx, col_idx)`.
    fn cell(&self, py: Python<'_>, row_idx: usize, col_idx: usize) -> PyResult<Cell> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc(&self.part)?;
        let node = doc
            .children_named(self.node, ns::A, "tr")
            .get(row_idx)
            .and_then(|&tr| doc.children_named(tr, ns::A, "tc").get(col_idx).copied())
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("cell index out of range"))?;
        drop(prs);
        Ok(Cell {
            prs: self.prs.clone_ref(py),
            part: self.part.clone(),
            node,
        })
    }

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

    #[setter(text)]
    fn set_text(&self, py: Python<'_>, value: &str) -> PyResult<()> {
        let mut prs = self.prs.borrow_mut(py);
        let doc = prs.doc_mut(&self.part)?;
        let tx_body = match doc.first_child_named(self.node, ns::A, "txBody") {
            Some(n) => n,
            None => {
                let tx_body = doc.create_element(ns::A, "a", "txBody", &[]);
                let body_pr = doc.create_element(ns::A, "a", "bodyPr", &[]);
                let lst_style = doc.create_element(ns::A, "a", "lstStyle", &[]);
                doc.append_child(tx_body, body_pr);
                doc.append_child(tx_body, lst_style);
                // a:tc schema orders txBody before tcPr
                doc.insert_child(self.node, 0, tx_body);
                tx_body
            }
        };
        set_tx_body_text(doc, tx_body, value);
        Ok(())
    }
}
