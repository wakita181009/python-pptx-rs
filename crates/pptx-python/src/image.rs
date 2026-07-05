//! Image read API for picture shapes.

use pyo3::exceptions::PyKeyError;
use pyo3::prelude::*;
use pyo3::types::PyBytes;

#[pyclass(name = "Image", module = "pptx_rs._core")]
pub struct Image {
    pub(crate) blob: Vec<u8>,
    /// Extension from the image part name (python-pptx sniffs the blob header
    /// instead; the part name is authoritative for well-formed packages).
    pub(crate) ext: String,
}

#[pymethods]
impl Image {
    #[getter]
    fn blob<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.blob)
    }

    #[getter]
    fn ext(&self) -> &str {
        &self.ext
    }

    #[getter]
    fn content_type(&self) -> PyResult<&'static str> {
        // python-pptx opc/spec.py image_content_types
        Ok(match self.ext.as_str() {
            "bmp" => "image/bmp",
            "emf" => "image/x-emf",
            "gif" => "image/gif",
            "jpe" | "jpeg" | "jpg" => "image/jpeg",
            "png" => "image/png",
            "tif" | "tiff" => "image/tiff",
            "wdp" => "image/vnd.ms-photo",
            "wmf" => "image/x-wmf",
            other => return Err(PyKeyError::new_err(other.to_string())),
        })
    }

    /// python-pptx returns the generic `image.<ext>` name for images loaded
    /// from a package (a real filename only exists for images added from a
    /// filesystem path at runtime).
    #[getter]
    fn filename(&self) -> String {
        format!("image.{}", self.ext)
    }
}
