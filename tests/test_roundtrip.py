"""Markdown written to a page and read back must mean the same thing, and reading must be stable."""

import pytest

from local_onenote_mcp.pagexml import build_page_update_xml
from local_onenote_mcp.reader import inline_text, render_page


def roundtrip(content: str, fmt: str = "markdown") -> str:
    xml = build_page_update_xml("p", content=content, content_format=fmt)
    return render_page(xml, "markdown", include_title=False)


def as_text(content: str, fmt: str = "markdown") -> str:
    return render_page(build_page_update_xml("p", content=content, content_format=fmt), "text", include_title=False)


SAMPLE = """# Plan

First paragraph with **bold**, *italic*, ~~gone~~ and `code`.

Second paragraph.

- one
- two
  - nested

1. step
2. step

Text after the list.

> quoted line

| A | B |
| --- | --- |
| x | **y** |

```
def f():
    return 1
```

Last line."""


def test_paragraphs_lists_and_blocks_survive_a_round_trip():
    once = roundtrip(SAMPLE)

    assert "First paragraph with **bold**, *italic*, ~~gone~~ and `code`.\n\nSecond paragraph." in once
    assert "- one\n- two\n  - nested" in once
    assert "2. step\n\nText after the list." in once
    assert "> quoted line\n\n| A | B |" in once
    assert roundtrip(once) == once


@pytest.mark.parametrize(
    "literal",
    ["- not a list", "# not a heading", "2*3*4", "<b>not html</b>", "1. not numbered", "a_b and _c_", "[x](y) link-like", "> not quote"],
)
def test_literal_text_is_escaped_so_it_stays_literal(literal):
    markdown = roundtrip(literal, "plain")

    assert as_text(markdown) == literal


def test_bold_survives_style_attributes_wrapped_by_onenote():
    assert inline_text("<span style='font-weight:\r\nbold'>x</span> y", markdown=True) == "**x** y"


def test_br_followed_by_a_raw_newline_is_one_line_break():
    assert inline_text("a<br />\r\nb") == "a\nb"
    assert inline_text("long line\r\nwrapped") == "long line wrapped"


def test_nested_emphasis_is_opened_and_closed_once():
    fragment = "<span style='font-style:italic'>(note: <span style='font-weight:bold'>mean of u</span>, ok)</span>"

    assert inline_text(fragment, markdown=True) == "*(note: **mean of u**, ok)*"


def test_code_with_backticks_gets_a_longer_fence():
    once = roundtrip("````\nuse ``` here\n````")

    assert once.startswith("````\nuse ``` here\n````")
    assert roundtrip(once) == once
