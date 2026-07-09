"""Type stubs for the Rust extension module."""

import enum
import os
from collections.abc import Iterator
from typing import IO

__version__: str

class MSO_SHAPE_TYPE(enum.IntEnum):
    AUTO_SHAPE = 1
    CALLOUT = 2
    CANVAS = 20
    CHART = 3
    COMMENT = 4
    DIAGRAM = 21
    EMBEDDED_OLE_OBJECT = 7
    FORM_CONTROL = 8
    FREEFORM = 5
    GROUP = 6
    IGX_GRAPHIC = 24
    INK = 22
    INK_COMMENT = 23
    LINE = 9
    LINKED_OLE_OBJECT = 10
    LINKED_PICTURE = 11
    MEDIA = 16
    OLE_CONTROL_OBJECT = 12
    PICTURE = 13
    PLACEHOLDER = 14
    SCRIPT_ANCHOR = 18
    TABLE = 19
    TEXT_BOX = 17
    TEXT_EFFECT = 15
    WEB_VIDEO = 26
    MIXED = -2

class Presentation:
    @staticmethod
    def from_bytes(data: bytes) -> Presentation: ...
    @staticmethod
    def from_path(path: str) -> Presentation: ...
    @property
    def slides(self) -> Slides: ...
    @property
    def slide_masters(self) -> SlideMasters: ...
    @property
    def slide_layouts(self) -> SlideLayouts: ...
    @property
    def slide_width(self) -> int | None: ...
    @property
    def slide_height(self) -> int | None: ...
    def save(self, target: str | os.PathLike[str] | IO[bytes]) -> None: ...

class Slides:
    def __len__(self) -> int: ...
    def __getitem__(self, idx: int) -> Slide: ...
    def __iter__(self) -> Iterator[Slide]: ...
    def add_slide(self, slide_layout: SlideLayout) -> Slide: ...

class Slide:
    @property
    def slide_id(self) -> int: ...
    @property
    def name(self) -> str: ...
    @property
    def shapes(self) -> SlideShapes: ...
    @property
    def slide_layout(self) -> SlideLayout: ...
    @property
    def has_notes_slide(self) -> bool: ...
    @property
    def notes_slide(self) -> NotesSlide: ...

class NotesSlide:
    @property
    def notes_text_frame(self) -> TextFrame | None: ...

class SlideMasters:
    def __len__(self) -> int: ...
    def __getitem__(self, idx: int) -> SlideMaster: ...
    def __iter__(self) -> Iterator[SlideMaster]: ...

class SlideMaster:
    @property
    def slide_layouts(self) -> SlideLayouts: ...

class SlideLayouts:
    def __len__(self) -> int: ...
    def __getitem__(self, idx: int) -> SlideLayout: ...
    def __iter__(self) -> Iterator[SlideLayout]: ...
    def get_by_name(self, name: str, default: object = None) -> SlideLayout | None: ...
    def index(self, slide_layout: SlideLayout) -> int: ...

class SlideLayout:
    @property
    def name(self) -> str: ...
    @property
    def slide_master(self) -> SlideMaster: ...

class MSO_AUTO_SHAPE_TYPE(enum.IntEnum):
    RECTANGLE = 1
    PARALLELOGRAM = 2
    TRAPEZOID = 3
    DIAMOND = 4
    ROUNDED_RECTANGLE = 5
    OCTAGON = 6
    ISOSCELES_TRIANGLE = 7
    RIGHT_TRIANGLE = 8
    OVAL = 9
    HEXAGON = 10
    CROSS = 11
    REGULAR_PENTAGON = 12
    CAN = 13
    CUBE = 14
    BEVEL = 15
    FOLDED_CORNER = 16
    SMILEY_FACE = 17
    DONUT = 18
    NO_SYMBOL = 19
    BLOCK_ARC = 20
    HEART = 21
    LIGHTNING_BOLT = 22
    SUN = 23
    MOON = 24
    ARC = 25
    DOUBLE_BRACKET = 26
    DOUBLE_BRACE = 27
    PLAQUE = 28
    LEFT_BRACKET = 29
    RIGHT_BRACKET = 30
    LEFT_BRACE = 31
    RIGHT_BRACE = 32
    RIGHT_ARROW = 33
    LEFT_ARROW = 34
    UP_ARROW = 35
    DOWN_ARROW = 36
    LEFT_RIGHT_ARROW = 37
    UP_DOWN_ARROW = 38
    QUAD_ARROW = 39
    LEFT_RIGHT_UP_ARROW = 40
    BENT_ARROW = 41
    U_TURN_ARROW = 42
    LEFT_UP_ARROW = 43
    BENT_UP_ARROW = 44
    CURVED_RIGHT_ARROW = 45
    CURVED_LEFT_ARROW = 46
    CURVED_UP_ARROW = 47
    CURVED_DOWN_ARROW = 48
    STRIPED_RIGHT_ARROW = 49
    NOTCHED_RIGHT_ARROW = 50
    PENTAGON = 51
    CHEVRON = 52
    RIGHT_ARROW_CALLOUT = 53
    LEFT_ARROW_CALLOUT = 54
    UP_ARROW_CALLOUT = 55
    DOWN_ARROW_CALLOUT = 56
    LEFT_RIGHT_ARROW_CALLOUT = 57
    UP_DOWN_ARROW_CALLOUT = 58
    QUAD_ARROW_CALLOUT = 59
    CIRCULAR_ARROW = 60
    FLOWCHART_PROCESS = 61
    FLOWCHART_ALTERNATE_PROCESS = 62
    FLOWCHART_DECISION = 63
    FLOWCHART_DATA = 64
    FLOWCHART_PREDEFINED_PROCESS = 65
    FLOWCHART_INTERNAL_STORAGE = 66
    FLOWCHART_DOCUMENT = 67
    FLOWCHART_MULTIDOCUMENT = 68
    FLOWCHART_TERMINATOR = 69
    FLOWCHART_PREPARATION = 70
    FLOWCHART_MANUAL_INPUT = 71
    FLOWCHART_MANUAL_OPERATION = 72
    FLOWCHART_CONNECTOR = 73
    FLOWCHART_OFFPAGE_CONNECTOR = 74
    FLOWCHART_CARD = 75
    FLOWCHART_PUNCHED_TAPE = 76
    FLOWCHART_SUMMING_JUNCTION = 77
    FLOWCHART_OR = 78
    FLOWCHART_COLLATE = 79
    FLOWCHART_SORT = 80
    FLOWCHART_EXTRACT = 81
    FLOWCHART_MERGE = 82
    FLOWCHART_STORED_DATA = 83
    FLOWCHART_DELAY = 84
    FLOWCHART_SEQUENTIAL_ACCESS_STORAGE = 85
    FLOWCHART_MAGNETIC_DISK = 86
    FLOWCHART_DIRECT_ACCESS_STORAGE = 87
    FLOWCHART_DISPLAY = 88
    EXPLOSION1 = 89
    EXPLOSION2 = 90
    STAR_4_POINT = 91
    STAR_5_POINT = 92
    STAR_8_POINT = 93
    STAR_16_POINT = 94
    STAR_24_POINT = 95
    STAR_32_POINT = 96
    UP_RIBBON = 97
    DOWN_RIBBON = 98
    CURVED_UP_RIBBON = 99
    CURVED_DOWN_RIBBON = 100
    VERTICAL_SCROLL = 101
    HORIZONTAL_SCROLL = 102
    WAVE = 103
    DOUBLE_WAVE = 104
    RECTANGULAR_CALLOUT = 105
    ROUNDED_RECTANGULAR_CALLOUT = 106
    OVAL_CALLOUT = 107
    CLOUD_CALLOUT = 108
    LINE_CALLOUT_1 = 109
    LINE_CALLOUT_2 = 110
    LINE_CALLOUT_3 = 111
    LINE_CALLOUT_4 = 112
    LINE_CALLOUT_1_ACCENT_BAR = 113
    LINE_CALLOUT_2_ACCENT_BAR = 114
    LINE_CALLOUT_3_ACCENT_BAR = 115
    LINE_CALLOUT_4_ACCENT_BAR = 116
    LINE_CALLOUT_1_NO_BORDER = 117
    LINE_CALLOUT_2_NO_BORDER = 118
    LINE_CALLOUT_3_NO_BORDER = 119
    LINE_CALLOUT_4_NO_BORDER = 120
    LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR = 121
    LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR = 122
    LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR = 123
    LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR = 124
    ACTION_BUTTON_CUSTOM = 125
    ACTION_BUTTON_HOME = 126
    ACTION_BUTTON_HELP = 127
    ACTION_BUTTON_INFORMATION = 128
    ACTION_BUTTON_BACK_OR_PREVIOUS = 129
    ACTION_BUTTON_FORWARD_OR_NEXT = 130
    ACTION_BUTTON_BEGINNING = 131
    ACTION_BUTTON_END = 132
    ACTION_BUTTON_RETURN = 133
    ACTION_BUTTON_DOCUMENT = 134
    ACTION_BUTTON_SOUND = 135
    ACTION_BUTTON_MOVIE = 136
    BALLOON = 137
    FLOWCHART_OFFLINE_STORAGE = 139
    LEFT_RIGHT_RIBBON = 140
    DIAGONAL_STRIPE = 141
    PIE = 142
    NON_ISOSCELES_TRAPEZOID = 143
    DECAGON = 144
    HEPTAGON = 145
    DODECAGON = 146
    STAR_6_POINT = 147
    STAR_7_POINT = 148
    STAR_10_POINT = 149
    STAR_12_POINT = 150
    ROUND_1_RECTANGLE = 151
    ROUND_2_SAME_RECTANGLE = 152
    ROUND_2_DIAG_RECTANGLE = 153
    SNIP_ROUND_RECTANGLE = 154
    SNIP_1_RECTANGLE = 155
    SNIP_2_SAME_RECTANGLE = 156
    SNIP_2_DIAG_RECTANGLE = 157
    FRAME = 158
    HALF_FRAME = 159
    TEAR = 160
    CHORD = 161
    CORNER = 162
    MATH_PLUS = 163
    MATH_MINUS = 164
    MATH_MULTIPLY = 165
    MATH_DIVIDE = 166
    MATH_EQUAL = 167
    MATH_NOT_EQUAL = 168
    CORNER_TABS = 169
    SQUARE_TABS = 170
    PLAQUE_TABS = 171
    GEAR_6 = 172
    GEAR_9 = 173
    FUNNEL = 174
    PIE_WEDGE = 175
    LEFT_CIRCULAR_ARROW = 176
    LEFT_RIGHT_CIRCULAR_ARROW = 177
    SWOOSH_ARROW = 178
    CLOUD = 179
    CHART_X = 180
    CHART_STAR = 181
    CHART_PLUS = 182
    LINE_INVERSE = 183

class SlideShapes:
    def __len__(self) -> int: ...
    def __getitem__(self, idx: int) -> Shape: ...
    def __iter__(self) -> Iterator[Shape]: ...
    def add_textbox(self, left: int, top: int, width: int, height: int) -> Shape: ...
    def add_picture(
        self,
        image_file: str | IO[bytes],
        left: int,
        top: int,
        width: int | None = None,
        height: int | None = None,
    ) -> Picture: ...
    def add_shape(self, autoshape_type_id: int, left: int, top: int, width: int, height: int) -> Shape: ...
    def add_table(self, rows: int, cols: int, left: int, top: int, width: int, height: int) -> GraphicFrame: ...
    @property
    def title(self) -> Shape | None: ...

class Shape:
    @property
    def shape_id(self) -> int: ...
    @property
    def name(self) -> str: ...
    @property
    def has_text_frame(self) -> bool: ...
    @property
    def text_frame(self) -> TextFrame: ...
    @property
    def shape_type(self) -> MSO_SHAPE_TYPE | None: ...
    @property
    def has_chart(self) -> bool: ...
    @property
    def chart(self) -> Chart: ...
    @property
    def table(self) -> Table: ...
    @property
    def image(self) -> Image: ...
    @property
    def shapes(self) -> SlideShapes: ...
    @property
    def _element(self) -> _XmlElement: ...
    text: str
    left: int | None
    top: int | None
    width: int | None
    height: int | None

class Picture(Shape): ...
class GraphicFrame(Shape): ...
class GroupShape(Shape): ...
class Connector(Shape): ...

class _XmlElement:
    def __getattr__(self, name: str) -> _XmlElement: ...
    @property
    def attrib(self) -> dict[str, str]: ...

class Table:
    def cell(self, row_idx: int, col_idx: int) -> _Cell: ...
    @property
    def rows(self) -> list[_Row]: ...

class _Row:
    @property
    def cells(self) -> list[_Cell]: ...

class _Cell:
    @property
    def text(self) -> str: ...
    @text.setter
    def text(self, value: str) -> None: ...

class Chart:
    @property
    def has_title(self) -> bool: ...
    @property
    def chart_title(self) -> ChartTitle: ...
    @property
    def plots(self) -> list[_BasePlot]: ...
    @property
    def series(self) -> list[Series]: ...

class ChartTitle:
    @property
    def text_frame(self) -> TextFrame: ...

class _BasePlot:
    @property
    def categories(self) -> list[Category]: ...

class Category:
    @property
    def label(self) -> str: ...

class Series:
    @property
    def name(self) -> str: ...
    @property
    def values(self) -> tuple[float | None, ...]: ...

class Image:
    @property
    def blob(self) -> bytes: ...
    @property
    def ext(self) -> str: ...
    @property
    def content_type(self) -> str: ...
    @property
    def filename(self) -> str: ...

class TextFrame:
    @property
    def paragraphs(self) -> list[_Paragraph]: ...
    text: str
    def add_paragraph(self) -> _Paragraph: ...

class _Paragraph:
    @property
    def runs(self) -> list[_Run]: ...
    text: str
    def add_run(self) -> _Run: ...

class _Run:
    text: str
