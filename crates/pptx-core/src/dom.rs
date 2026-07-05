//! Owned, arena-based XML DOM with namespace resolution.
//!
//! Nodes are appended to a `Vec` arena and never deallocated individually, so a
//! `NodeId` handed out to a caller (e.g. a Python proxy object) stays valid for
//! the life of the document even across detach/insert operations.

use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};

pub type NodeId = usize;

#[derive(Debug, Clone, PartialEq)]
pub enum Child {
    Element(NodeId),
    Text(String),
}

#[derive(Debug, Clone)]
pub struct Element {
    pub ns: Option<String>,
    pub prefix: Option<String>,
    pub local: String,
    /// Attributes with raw qualified names (including `xmlns:*` declarations),
    /// in document order. Values are unescaped.
    pub attrs: Vec<(String, String)>,
    pub children: Vec<Child>,
    pub parent: Option<NodeId>,
}

#[derive(Debug, Clone)]
pub struct Document {
    elems: Vec<Element>,
    pub root: NodeId,
}

impl Document {
    pub fn parse(bytes: &[u8]) -> Result<Document> {
        let bytes = bytes.strip_prefix(b"\xef\xbb\xbf").unwrap_or(bytes);
        let text = std::str::from_utf8(bytes).map_err(|e| Error::Xml(e.to_string()))?;
        let mut reader = NsReader::from_str(text);
        reader.config_mut().expand_empty_elements = true;

        let mut elems: Vec<Element> = Vec::new();
        let mut stack: Vec<NodeId> = Vec::new();
        let mut root: Option<NodeId> = None;

        loop {
            let (resolve, event) = reader
                .read_resolved_event()
                .map_err(|e| Error::Xml(e.to_string()))?;
            match event {
                Event::Start(e) => {
                    let ns = match resolve {
                        ResolveResult::Bound(ns) => {
                            Some(String::from_utf8_lossy(ns.as_ref()).into_owned())
                        }
                        _ => None,
                    };
                    let name = e.name();
                    let local = String::from_utf8_lossy(name.local_name().as_ref()).into_owned();
                    let prefix = name
                        .prefix()
                        .map(|p| String::from_utf8_lossy(p.as_ref()).into_owned());
                    let mut attrs = Vec::new();
                    for attr in e.attributes() {
                        let attr = attr.map_err(|e| Error::Xml(e.to_string()))?;
                        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                        let value = attr
                            .normalized_value(quick_xml::XmlVersion::Explicit1_0)
                            .map_err(|e| Error::Xml(e.to_string()))?
                            .into_owned();
                        attrs.push((key, value));
                    }
                    let id = elems.len();
                    elems.push(Element {
                        ns,
                        prefix,
                        local,
                        attrs,
                        children: Vec::new(),
                        parent: stack.last().copied(),
                    });
                    if let Some(&parent) = stack.last() {
                        elems[parent].children.push(Child::Element(id));
                    } else if root.is_none() {
                        root = Some(id);
                    }
                    stack.push(id);
                }
                Event::End(_) => {
                    stack.pop();
                }
                Event::Text(t) => {
                    if let Some(&parent) = stack.last() {
                        let s = t
                            .xml_content(quick_xml::XmlVersion::Explicit1_0)
                            .map_err(|e| Error::Xml(e.to_string()))?;
                        if !s.is_empty()
                            && (!s.trim().is_empty()
                                || keeps_whitespace(&elems[parent])
                                || has_trailing_text(&elems[parent]))
                        {
                            push_text(&mut elems[parent], &s);
                        }
                    }
                }
                Event::GeneralRef(r) => {
                    if let Some(&parent) = stack.last() {
                        let name = r.decode().map_err(|e| Error::Xml(e.to_string()))?;
                        let resolved = resolve_reference(&name)?;
                        push_text(&mut elems[parent], &resolved);
                    }
                }
                Event::CData(t) => {
                    if let Some(&parent) = stack.last() {
                        let s = String::from_utf8_lossy(&t).into_owned();
                        push_text(&mut elems[parent], &s);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let root = root.ok_or_else(|| Error::Xml("document has no root element".into()))?;
        Ok(Document { elems, root })
    }

    pub fn serialize(&self) -> Vec<u8> {
        let mut out = String::with_capacity(4096);
        out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
        self.write_element(self.root, &mut out);
        out.into_bytes()
    }

    fn write_element(&self, id: NodeId, out: &mut String) {
        let el = &self.elems[id];
        out.push('<');
        if let Some(p) = &el.prefix {
            out.push_str(p);
            out.push(':');
        }
        out.push_str(&el.local);
        for (k, v) in &el.attrs {
            out.push(' ');
            out.push_str(k);
            out.push_str("=\"");
            escape_into(v, out, true);
            out.push('"');
        }
        if el.children.is_empty() {
            out.push_str("/>");
            return;
        }
        out.push('>');
        for child in &el.children {
            match child {
                Child::Element(c) => self.write_element(*c, out),
                Child::Text(t) => escape_into(t, out, false),
            }
        }
        out.push_str("</");
        if let Some(p) = &el.prefix {
            out.push_str(p);
            out.push(':');
        }
        out.push_str(&el.local);
        out.push('>');
    }

    // -- node accessors -------------------------------------------------

    pub fn get(&self, id: NodeId) -> &Element {
        &self.elems[id]
    }

    pub fn get_mut(&mut self, id: NodeId) -> &mut Element {
        &mut self.elems[id]
    }

    pub fn is(&self, id: NodeId, ns: &str, local: &str) -> bool {
        let el = &self.elems[id];
        el.local == local && el.ns.as_deref() == Some(ns)
    }

    pub fn attr(&self, id: NodeId, name: &str) -> Option<&str> {
        self.elems[id]
            .attrs
            .iter()
            .find(|(k, _)| k == name)
            .map(|(_, v)| v.as_str())
    }

    pub fn set_attr(&mut self, id: NodeId, name: &str, value: &str) {
        let el = &mut self.elems[id];
        match el.attrs.iter_mut().find(|(k, _)| k == name) {
            Some((_, v)) => *v = value.to_string(),
            None => el.attrs.push((name.to_string(), value.to_string())),
        }
    }

    /// Element children ids, in order.
    pub fn child_elements(&self, id: NodeId) -> Vec<NodeId> {
        self.elems[id]
            .children
            .iter()
            .filter_map(|c| match c {
                Child::Element(e) => Some(*e),
                Child::Text(_) => None,
            })
            .collect()
    }

    pub fn children_named(&self, id: NodeId, ns: &str, local: &str) -> Vec<NodeId> {
        self.child_elements(id)
            .into_iter()
            .filter(|&c| self.is(c, ns, local))
            .collect()
    }

    pub fn first_child_named(&self, id: NodeId, ns: &str, local: &str) -> Option<NodeId> {
        self.child_elements(id)
            .into_iter()
            .find(|&c| self.is(c, ns, local))
    }

    /// Concatenated direct text content (as for `<a:t>`).
    pub fn text(&self, id: NodeId) -> String {
        let mut s = String::new();
        for child in &self.elems[id].children {
            if let Child::Text(t) = child {
                s.push_str(t);
            }
        }
        s
    }

    pub fn set_text(&mut self, id: NodeId, text: &str) {
        let el = &mut self.elems[id];
        el.children.retain(|c| matches!(c, Child::Element(_)));
        if !text.is_empty() {
            el.children.push(Child::Text(text.to_string()));
        }
    }

    /// Walk the subtree rooted at `id` in document order, invoking `f` on each element.
    pub fn walk(&self, id: NodeId, f: &mut impl FnMut(&Document, NodeId)) {
        f(self, id);
        for child in self.child_elements(id) {
            self.walk(child, f);
        }
    }

    // -- mutation --------------------------------------------------------

    pub fn create_element(
        &mut self,
        ns: &str,
        prefix: &str,
        local: &str,
        attrs: &[(&str, &str)],
    ) -> NodeId {
        let id = self.elems.len();
        self.elems.push(Element {
            ns: Some(ns.to_string()),
            prefix: if prefix.is_empty() {
                None
            } else {
                Some(prefix.to_string())
            },
            local: local.to_string(),
            attrs: attrs
                .iter()
                .map(|(k, v)| (k.to_string(), v.to_string()))
                .collect(),
            children: Vec::new(),
            parent: None,
        });
        id
    }

    pub fn append_child(&mut self, parent: NodeId, child: NodeId) {
        self.elems[child].parent = Some(parent);
        self.elems[parent].children.push(Child::Element(child));
    }

    pub fn insert_child(&mut self, parent: NodeId, index: usize, child: NodeId) {
        self.elems[child].parent = Some(parent);
        self.elems[parent]
            .children
            .insert(index, Child::Element(child));
    }

    /// Detach `child` from `parent`. The node stays allocated so existing
    /// `NodeId`s never dangle.
    pub fn remove_child(&mut self, parent: NodeId, child: NodeId) {
        self.elems[parent]
            .children
            .retain(|c| !matches!(c, Child::Element(e) if *e == child));
        self.elems[child].parent = None;
    }
}

/// Whitespace-only text is significant only inside `<a:t>` runs.
fn keeps_whitespace(el: &Element) -> bool {
    el.local == "t" && el.ns.as_deref() == Some(crate::ns::A)
}

/// True when the element's last child is text (e.g. the segment before an
/// entity reference), so a following whitespace-only segment belongs to it.
fn has_trailing_text(el: &Element) -> bool {
    matches!(el.children.last(), Some(Child::Text(_)))
}

/// Append text, merging with a trailing text child so entity references don't
/// fragment character data.
fn push_text(el: &mut Element, s: &str) {
    if let Some(Child::Text(t)) = el.children.last_mut() {
        t.push_str(s);
    } else {
        el.children.push(Child::Text(s.to_string()));
    }
}

/// Resolve a general reference name (`amp`, `#38`, `#x26`, ...) to its text.
fn resolve_reference(name: &str) -> Result<String> {
    if let Some(num) = name.strip_prefix('#') {
        let code = match num.strip_prefix(['x', 'X']) {
            Some(hex) => u32::from_str_radix(hex, 16),
            None => num.parse(),
        }
        .map_err(|_| Error::Xml(format!("invalid character reference &{name};")))?;
        let c = char::from_u32(code)
            .ok_or_else(|| Error::Xml(format!("invalid character reference &{name};")))?;
        return Ok(c.to_string());
    }
    quick_xml::escape::resolve_predefined_entity(name)
        .map(str::to_string)
        .ok_or_else(|| Error::Xml(format!("unsupported entity &{name};")))
}

fn escape_into(s: &str, out: &mut String, is_attr: bool) {
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' if is_attr => out.push_str("&quot;"),
            _ => out.push(c),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SLIDE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Hello &amp; goodbye</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>"#;

    #[test]
    fn parse_resolves_namespaces_and_unescapes() {
        let doc = Document::parse(SLIDE.as_bytes()).unwrap();
        assert!(doc.is(doc.root, crate::ns::P, "sld"));
        let mut texts = Vec::new();
        doc.walk(doc.root, &mut |d, id| {
            if d.is(id, crate::ns::A, "t") {
                texts.push(d.text(id));
            }
        });
        assert_eq!(texts, vec!["Hello & goodbye"]);
    }

    #[test]
    fn round_trip_preserves_content() {
        let doc = Document::parse(SLIDE.as_bytes()).unwrap();
        let out = doc.serialize();
        let doc2 = Document::parse(&out).unwrap();
        let mut texts = Vec::new();
        doc2.walk(doc2.root, &mut |d, id| {
            if d.is(id, crate::ns::A, "t") {
                texts.push(d.text(id));
            }
        });
        assert_eq!(texts, vec!["Hello & goodbye"]);
    }

    #[test]
    fn mutation_and_stable_ids() {
        let mut doc = Document::parse(SLIDE.as_bytes()).unwrap();
        let mut t_id = None;
        doc.walk(doc.root, &mut |d, id| {
            if d.is(id, crate::ns::A, "t") {
                t_id = Some(id);
            }
        });
        let t = t_id.unwrap();
        doc.set_text(t, "replaced <text>");
        let out = String::from_utf8(doc.serialize()).unwrap();
        assert!(out.contains("replaced &lt;text&gt;"));
    }
}
