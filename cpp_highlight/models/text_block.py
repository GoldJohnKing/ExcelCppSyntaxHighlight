"""Custom TextBlock for whitespace preservation."""

from xml.etree.ElementTree import Element

from openpyxl.cell.rich_text import TextBlock as _OrigTextBlock


class TextBlock(_OrigTextBlock):
    """TextBlock that properly preserves whitespace in Excel."""

    def to_tree(self):
        """Convert to XML tree, ensuring whitespace is preserved."""
        el = Element("r")
        if self.font:
            el.append(self.font.to_tree(tagname="rPr"))
        t = Element("t")
        t.text = self.text

        # Always set xml:space="preserve" if text contains any whitespace
        if self.text and (self.text != self.text.strip() or not self.text.strip()):
            XML_NS = "http://www.w3.org/XML/1998/namespace"
            t.set("{%s}space" % XML_NS, "preserve")

        el.append(t)
        return el
