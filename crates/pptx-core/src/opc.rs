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
