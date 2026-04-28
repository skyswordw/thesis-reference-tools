from pathlib import Path
from xml.etree import ElementTree


CSL_NS = "http://purl.org/net/xbiblio/csl"


def test_neu_thesis_csl_uses_single_space_after_reference_number_without_flush_alignment():
    root = ElementTree.parse(Path("styles/neu-thesis-gbt7714-2015-numeric.csl")).getroot()
    ns = {"csl": CSL_NS}
    bibliography = root.find("csl:bibliography", ns)
    assert bibliography is not None
    assert "second-field-align" not in bibliography.attrib

    citation_numbers = bibliography.findall("./csl:layout/csl:text[@variable='citation-number']", ns)
    assert citation_numbers
    for citation_number in citation_numbers:
        assert citation_number.attrib["prefix"] == "["
        assert citation_number.attrib["suffix"] == "] "


def test_neu_thesis_csl_has_english_bibliography_locale_for_et_al_without_losing_chinese_default():
    root = ElementTree.parse(Path("styles/neu-thesis-gbt7714-2015-numeric.csl")).getroot()
    ns = {"csl": CSL_NS}

    bibliography = root.find("csl:bibliography", ns)
    assert bibliography is not None
    layouts = bibliography.findall("csl:layout", ns)
    layout_locales = [layout.attrib.get("locale") for layout in layouts]

    assert "en" in layout_locales
    assert None in layout_locales
    assert layout_locales != [None]

    english_locale = root.find("csl:locale[@xml:lang='en']", {**ns, "xml": "http://www.w3.org/XML/1998/namespace"})
    assert english_locale is not None
    english_et_al = english_locale.find("./csl:terms/csl:term[@name='et-al']", ns)
    assert english_et_al is not None
    assert english_et_al.text == "et al."
