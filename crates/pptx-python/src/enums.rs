//! python-pptx enum compatibility (`pptx.enum.*`).
//!
//! Each enum is exposed as a real Python `enum.IntEnum` subclass via `pyenum`,
//! with member names and values copied verbatim from python-pptx's
//! `pptx/enum/*.py` so `==` against real python-pptx members works through the
//! int mixin. pyenum uses the Rust variant identifiers as Python member names,
//! hence the SCREAMING_SNAKE variants.

use pyenum::PyEnum;

/// `pptx.enum.shapes.MSO_SHAPE_TYPE`
#[allow(non_camel_case_types, clippy::upper_case_acronyms)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, PyEnum)]
#[pyenum(base = "IntEnum")]
pub enum MSO_SHAPE_TYPE {
    AUTO_SHAPE = 1,
    CALLOUT = 2,
    CANVAS = 20,
    CHART = 3,
    COMMENT = 4,
    DIAGRAM = 21,
    EMBEDDED_OLE_OBJECT = 7,
    FORM_CONTROL = 8,
    FREEFORM = 5,
    GROUP = 6,
    IGX_GRAPHIC = 24,
    INK = 22,
    INK_COMMENT = 23,
    LINE = 9,
    LINKED_OLE_OBJECT = 10,
    LINKED_PICTURE = 11,
    MEDIA = 16,
    OLE_CONTROL_OBJECT = 12,
    PICTURE = 13,
    PLACEHOLDER = 14,
    SCRIPT_ANCHOR = 18,
    TABLE = 19,
    TEXT_BOX = 17,
    TEXT_EFFECT = 15,
    WEB_VIDEO = 26,
    MIXED = -2,
}
