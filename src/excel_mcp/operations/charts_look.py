"""The look of a chart made in current Excel: gray text, light gridlines, no frame on the plot."""

# pyright: reportArgumentType=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
# openpyxl's stubs type its enum-like fields as Literals and its descriptors as plain attributes.

from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.colors import ColorChoice, SchemeColor
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Font,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
    RichTextProperties,
)

_ACCENTS = 6
_TIERS = ((None, None), (60000, None), (80000, 20000))


def scheme_color(
    name: str, mod: int | None = None, off: int | None = None, alpha: int | None = None
) -> ColorChoice:
    return ColorChoice(schemeClr=SchemeColor(val=name, lumMod=mod, lumOff=off, alpha=alpha))


def theme_color(index: int, alpha: int | None = None) -> ColorChoice:
    """The automatic color of the nth series, as Excel cycles through the theme accents."""
    mod, off = _TIERS[index // _ACCENTS % len(_TIERS)]
    return scheme_color(f"accent{index % _ACCENTS + 1}", mod, off, alpha)


def thin_line(mod: int, off: int) -> LineProperties:
    return LineProperties(
        w=9525,
        cap="flat",
        cmpd="sng",
        algn="ctr",
        solidFill=scheme_color("tx1", mod, off),
        round=True,
    )


def gridlines(minor: bool = False) -> ChartLines:
    line = thin_line(5000, 95000) if minor else thin_line(15000, 85000)
    return ChartLines(spPr=GraphicalProperties(ln=line))


def no_fill() -> GraphicalProperties:
    return GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))


def text_properties(size: int, rotation: int | None = None) -> RichText:
    """Text style of titles, axes and legends: gray, the theme font, `size` in 1/100 pt."""
    properties = CharacterProperties(
        sz=size,
        b=False,
        i=False,
        u="none",
        strike="noStrike",
        kern=1200,
        baseline=0,
        solidFill=scheme_color("tx1", 65000, 35000),
        latin=Font(typeface="+mn-lt"),
        ea=Font(typeface="+mn-ea"),
        cs=Font(typeface="+mn-cs"),
    )
    body = RichTextProperties(
        rot=rotation,
        spcFirstLastPara=True,
        vertOverflow="ellipsis",
        vert="horz",
        wrap="square",
        anchor="ctr",
        anchorCtr=True,
    )
    paragraph = Paragraph(
        pPr=ParagraphProperties(defRPr=properties),
        r=[],
        endParaRPr=CharacterProperties(lang="en-US"),
    )
    return RichText(bodyPr=body, p=[paragraph])


def title(text: str, size: int, rotation: int | None = None) -> Title:
    """A title that sits beside the plot (Excel overlays it without the explicit flag)."""
    rich = text_properties(size, rotation)
    rich.p = [
        Paragraph(
            pPr=rich.p[0].pPr,
            r=[RegularTextRun(rPr=CharacterProperties(lang="en-US"), t=text)],
        )
    ]
    return Title(tx=Text(rich=rich), overlay=False, spPr=no_fill())
