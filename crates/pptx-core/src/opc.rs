//! OPC (Open Packaging Conventions) package: the zip container.
//!
//! Parts parse lazily; a part never touched keeps its original bytes and is
//! written back verbatim on save, so unsupported features round-trip intact.

use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::dom::Document;
use crate::error::{Error, Result};
use crate::ns;

pub struct Part {
    pub raw: Option<Vec<u8>>,
    pub doc: Option<Document>,
    pub dirty: bool,
}

pub struct Package {
    /// Part name (no leading slash) -> part, preserving original zip order.
    order: Vec<String>,
    parts: HashMap<String, Part>,
}

#[derive(Debug, Clone)]
pub struct Relationship {
    pub id: String,
    pub reltype: String,
    pub target: String,
    pub is_external: bool,
}

impl Package {
    pub fn from_bytes(bytes: &[u8]) -> Result<Package> {
        let mut archive = ZipArchive::new(Cursor::new(bytes))?;
        let mut order = Vec::with_capacity(archive.len());
        let mut parts = HashMap::with_capacity(archive.len());
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            if file.is_dir() {
                continue;
            }
            let name = file.name().to_string();
            let mut data = Vec::with_capacity(file.size() as usize);
            file.read_to_end(&mut data)?;
            order.push(name.clone());
            parts.insert(
                name,
                Part {
                    raw: Some(data),
                    doc: None,
                    dirty: false,
                },
            );
        }
        if !parts.contains_key("[Content_Types].xml") {
            return Err(Error::InvalidPackage("missing [Content_Types].xml".into()));
        }
        Ok(Package { order, parts })
    }

    pub fn part_names(&self) -> &[String] {
        &self.order
    }

    pub fn contains(&self, name: &str) -> bool {
        self.parts.contains_key(name)
    }

    pub fn raw(&self, name: &str) -> Option<&[u8]> {
        self.parts.get(name).and_then(|p| p.raw.as_deref())
    }

    /// Parse (once) and return the part's document.
    pub fn doc(&mut self, name: &str) -> Result<&Document> {
        self.ensure_parsed(name)?;
        Ok(self.parts[name].doc.as_ref().unwrap())
    }

    /// Parse (once), mark dirty, and return the part's document mutably.
    pub fn doc_mut(&mut self, name: &str) -> Result<&mut Document> {
        self.ensure_parsed(name)?;
        let part = self.parts.get_mut(name).unwrap();
        part.dirty = true;
        Ok(part.doc.as_mut().unwrap())
    }

    fn ensure_parsed(&mut self, name: &str) -> Result<()> {
        let part = self
            .parts
            .get_mut(name)
            .ok_or_else(|| Error::PartNotFound(name.to_string()))?;
        if part.doc.is_none() {
            let raw = part
                .raw
                .as_ref()
                .expect("part has neither raw bytes nor doc");
            part.doc = Some(Document::parse(raw)?);
        }
        Ok(())
    }

    pub fn save_to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for name in &self.order {
            let part = &self.parts[name];
            writer.start_file(name.as_str(), options)?;
            if part.dirty {
                writer.write_all(&part.doc.as_ref().unwrap().serialize())?;
            } else {
                writer.write_all(part.raw.as_ref().expect("clean part must have raw bytes"))?;
            }
        }
        Ok(writer.finish()?.into_inner())
    }

    // -- relationships ---------------------------------------------------

    /// Relationships of `part_name`, or an empty list when it has no rels part.
    pub fn rels(&mut self, part_name: &str) -> Result<Vec<Relationship>> {
        let rels_name = rels_part_name(part_name);
        if !self.contains(&rels_name) {
            return Ok(Vec::new());
        }
        let doc = self.doc(&rels_name)?;
        let mut rels = Vec::new();
        for child in doc.children_named(doc.root, ns::REL, "Relationship") {
            rels.push(Relationship {
                id: doc.attr(child, "Id").unwrap_or_default().to_string(),
                reltype: doc.attr(child, "Type").unwrap_or_default().to_string(),
                target: doc.attr(child, "Target").unwrap_or_default().to_string(),
                is_external: doc.attr(child, "TargetMode") == Some("External"),
            });
        }
        Ok(rels)
    }

    /// Absolute part name a relationship of `source_part` points at.
    pub fn resolve_target(&self, source_part: &str, target: &str) -> String {
        resolve_target(source_part, target)
    }

    /// Add a new XML part and register its content type as an Override.
    pub fn add_xml_part(&mut self, name: &str, content_type: &str, doc: Document) -> Result<()> {
        if self.contains(name) {
            return Err(Error::InvalidPackage(format!(
                "part already exists: {name}"
            )));
        }
        let ct = self.doc_mut(CONTENT_TYPES_PART)?;
        let root = ct.root;
        let override_el = ct.create_element(
            ns::CT,
            "",
            "Override",
            &[
                ("PartName", &format!("/{name}")),
                ("ContentType", content_type),
            ],
        );
        ct.append_child(root, override_el);
        self.order.push(name.to_string());
        self.parts.insert(
            name.to_string(),
            Part {
                raw: None,
                doc: Some(doc),
                dirty: true,
            },
        );
        Ok(())
    }

    /// Add a new binary (non-XML) part. The caller is responsible for making
    /// sure its extension is covered by a `Default` content-type rule.
    pub fn add_binary_part(&mut self, name: &str, blob: Vec<u8>) -> Result<()> {
        if self.contains(name) {
            return Err(Error::InvalidPackage(format!(
                "part already exists: {name}"
            )));
        }
        self.order.push(name.to_string());
        self.parts.insert(
            name.to_string(),
            Part {
                raw: Some(blob),
                doc: None,
                dirty: false,
            },
        );
        Ok(())
    }

    /// Register a `Default` content-type rule for `ext` unless one exists
    /// (extension comparison is case-insensitive per OPC).
    pub fn add_default_content_type(&mut self, ext: &str, content_type: &str) -> Result<()> {
        let ct = self.doc(CONTENT_TYPES_PART)?;
        let exists = ct
            .children_named(ct.root, ns::CT, "Default")
            .into_iter()
            .any(|d| {
                ct.attr(d, "Extension")
                    .is_some_and(|e| e.eq_ignore_ascii_case(ext))
            });
        if exists {
            return Ok(());
        }
        let ct = self.doc_mut(CONTENT_TYPES_PART)?;
        let root = ct.root;
        // schema requires Default elements before Override elements
        let index = ct
            .child_elements(root)
            .into_iter()
            .position(|c| ct.is(c, ns::CT, "Override"))
            .unwrap_or_else(|| ct.child_elements(root).len());
        let default_el = ct.create_element(
            ns::CT,
            "",
            "Default",
            &[("Extension", ext), ("ContentType", content_type)],
        );
        ct.insert_child(root, index, default_el);
        Ok(())
    }

    /// Return the rId of an existing relationship of `source_part` matching
    /// `reltype` and `target`, or add one (python-pptx `relate_to`).
    pub fn get_or_add_relationship(
        &mut self,
        source_part: &str,
        reltype: &str,
        target: &str,
    ) -> Result<String> {
        let existing = self
            .rels(source_part)?
            .into_iter()
            .find(|r| r.reltype == reltype && r.target == target && !r.is_external);
        match existing {
            Some(rel) => Ok(rel.id),
            None => self.add_relationship(source_part, reltype, target),
        }
    }

    /// Add a relationship from `source_part`, creating its rels part when
    /// missing, and return the assigned rId.
    pub fn add_relationship(
        &mut self,
        source_part: &str,
        reltype: &str,
        target: &str,
    ) -> Result<String> {
        let rels_name = rels_part_name(source_part);
        if !self.contains(&rels_name) {
            // .rels parts are covered by the package's Default content-type
            // rule, so no [Content_Types].xml entry is needed.
            let doc = Document::parse(EMPTY_RELS_XML.as_bytes())?;
            self.order.push(rels_name.clone());
            self.parts.insert(
                rels_name.clone(),
                Part {
                    raw: None,
                    doc: Some(doc),
                    dirty: true,
                },
            );
        }
        let doc = self.doc_mut(&rels_name)?;
        let max_n = doc
            .children_named(doc.root, ns::REL, "Relationship")
            .into_iter()
            .filter_map(|rel| {
                doc.attr(rel, "Id")
                    .and_then(|v| v.strip_prefix("rId"))
                    .and_then(|v| v.parse::<u32>().ok())
            })
            .max()
            .unwrap_or(0);
        let rid = format!("rId{}", max_n + 1);
        let root = doc.root;
        let el = doc.create_element(
            ns::REL,
            "",
            "Relationship",
            &[("Id", &rid), ("Type", reltype), ("Target", target)],
        );
        doc.append_child(root, el);
        Ok(rid)
    }
}

const CONTENT_TYPES_PART: &str = "[Content_Types].xml";

const EMPTY_RELS_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    "\r\n",
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
);

/// Relative path from `source_part`'s directory to `target_part`
/// (e.g. `ppt/slides/slide1.xml` → `ppt/slideLayouts/x.xml` is
/// `../slideLayouts/x.xml`).
pub fn relative_target(source_part: &str, target_part: &str) -> String {
    let src_dir: Vec<&str> = match source_part.rsplit_once('/') {
        Some((dir, _)) => dir.split('/').collect(),
        None => Vec::new(),
    };
    let tgt: Vec<&str> = target_part.split('/').collect();
    let (tgt_dir, tgt_base) = tgt.split_at(tgt.len() - 1);
    let common = src_dir
        .iter()
        .zip(tgt_dir.iter())
        .take_while(|(a, b)| a == b)
        .count();
    let mut segments: Vec<&str> = Vec::new();
    segments.extend(std::iter::repeat_n("..", src_dir.len() - common));
    segments.extend(&tgt_dir[common..]);
    segments.push(tgt_base[0]);
    segments.join("/")
}

pub fn rels_part_name(part_name: &str) -> String {
    match part_name.rsplit_once('/') {
        Some((dir, base)) => format!("{dir}/_rels/{base}.rels"),
        None => format!("_rels/{part_name}.rels"),
    }
}

/// Resolve a relationship target relative to its source part, normalizing `..`.
pub fn resolve_target(source_part: &str, target: &str) -> String {
    if let Some(abs) = target.strip_prefix('/') {
        return abs.to_string();
    }
    let mut segments: Vec<&str> = match source_part.rsplit_once('/') {
        Some((dir, _)) => dir.split('/').collect(),
        None => Vec::new(),
    };
    for seg in target.split('/') {
        match seg {
            "." | "" => {}
            ".." => {
                segments.pop();
            }
            s => segments.push(s),
        }
    }
    segments.join("/")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn minimal_package() -> Package {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        let options = SimpleFileOptions::default();
        writer.start_file(CONTENT_TYPES_PART, options).unwrap();
        writer
            .write_all(
                concat!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
                    r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
                    r#"<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>"#,
                    r#"</Types>"#,
                )
                .as_bytes(),
            )
            .unwrap();
        let bytes = writer.finish().unwrap().into_inner();
        Package::from_bytes(&bytes).unwrap()
    }

    #[test]
    fn adds_binary_part_and_roundtrips() {
        let mut pkg = minimal_package();
        pkg.add_binary_part("ppt/media/image1.png", vec![1, 2, 3])
            .unwrap();
        assert!(
            pkg.add_binary_part("ppt/media/image1.png", vec![9])
                .is_err()
        );
        let saved = pkg.save_to_bytes().unwrap();
        let pkg2 = Package::from_bytes(&saved).unwrap();
        assert_eq!(pkg2.raw("ppt/media/image1.png"), Some(&[1u8, 2, 3][..]));
    }

    #[test]
    fn default_content_type_added_once_before_overrides() {
        let mut pkg = minimal_package();
        pkg.add_default_content_type("png", "image/png").unwrap();
        pkg.add_default_content_type("PNG", "image/png").unwrap();
        pkg.add_default_content_type("rels", "ignored").unwrap();
        let doc = pkg.doc(CONTENT_TYPES_PART).unwrap();
        let defaults = doc.children_named(doc.root, ns::CT, "Default");
        assert_eq!(defaults.len(), 2);
        let children = doc.child_elements(doc.root);
        // new Default sits before the Override element
        assert!(doc.is(children[1], ns::CT, "Default"));
        assert!(doc.is(children[2], ns::CT, "Override"));
    }

    #[test]
    fn get_or_add_relationship_reuses_matching_rel() {
        let mut pkg = minimal_package();
        let r1 = pkg
            .get_or_add_relationship(
                "ppt/slides/slide1.xml",
                "http://reltype/image",
                "../media/image1.png",
            )
            .unwrap();
        let r2 = pkg
            .get_or_add_relationship(
                "ppt/slides/slide1.xml",
                "http://reltype/image",
                "../media/image1.png",
            )
            .unwrap();
        let r3 = pkg
            .get_or_add_relationship(
                "ppt/slides/slide1.xml",
                "http://reltype/image",
                "../media/image2.png",
            )
            .unwrap();
        assert_eq!(r1, r2);
        assert_ne!(r1, r3);
    }

    #[test]
    fn resolves_relative_targets() {
        assert_eq!(
            resolve_target("ppt/presentation.xml", "slides/slide1.xml"),
            "ppt/slides/slide1.xml"
        );
        assert_eq!(
            resolve_target("ppt/slides/slide1.xml", "../slideLayouts/slideLayout1.xml"),
            "ppt/slideLayouts/slideLayout1.xml"
        );
        assert_eq!(
            resolve_target("ppt/presentation.xml", "/docProps/core.xml"),
            "docProps/core.xml"
        );
    }

    #[test]
    fn computes_relative_targets() {
        assert_eq!(
            relative_target("ppt/presentation.xml", "ppt/slides/slide1.xml"),
            "slides/slide1.xml"
        );
        assert_eq!(
            relative_target("ppt/slides/slide1.xml", "ppt/slideLayouts/slideLayout2.xml"),
            "../slideLayouts/slideLayout2.xml"
        );
        assert_eq!(
            relative_target("", "docProps/core.xml"),
            "docProps/core.xml"
        );
    }

    #[test]
    fn rels_part_names() {
        assert_eq!(
            rels_part_name("ppt/presentation.xml"),
            "ppt/_rels/presentation.xml.rels"
        );
        assert_eq!(
            rels_part_name("[Content_Types].xml"),
            "_rels/[Content_Types].xml.rels"
        );
    }
}
