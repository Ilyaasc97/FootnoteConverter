# -*- coding: utf-8 -*-
"""
Self-test for footnoteconverter.py
==================================

Runs the converter logic against a light-weight simulation of the LibreOffice
UNO text API, so that the conversion, numbering, alignment and font
unification flows can be verified from the command line without starting
LibreOffice Writer.

Run it with the Python interpreter that ships with LibreOffice, for example:

    "C:\\Program Files\\LibreOffice\\program\\python.exe" tests\\selftest.py

The simulation implements the parts of the UNO API that the extension uses and
mirrors the Writer behaviour that the code relies on:

  * a footnote anchor occupies one character position, so anchor.getStart()
    and anchor.getEnd() bracket the footnote number,
  * text ranges and cursors are live objects that follow text edits
    (they shift, clamp or stay in place when the surrounding text changes),
  * inserting text exactly at a footnote anchor puts the text on the anchor's
    left side, which is what the "(" of a footnote reference is supposed to be.
"""

import importlib.util
import os
import re
import sys
import types

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXTENSION = os.path.join(ROOT, "footnoteconverter.py")


# ---------------------------------------------------------------------------
# 1. Stand-ins for the UNO modules that the extension imports
# ---------------------------------------------------------------------------
def _module(name, parent=None):
    mod = types.ModuleType(name)
    sys.modules[name] = mod
    if parent:
        setattr(sys.modules[parent], name.rsplit(".", 1)[1], mod)
    return mod


class _Base(object):
    pass


class _ImplementationHelper(object):
    def __init__(self):
        self.registered = []

    def addImplementation(self, class_, name, services):
        self.registered.append((class_, name, services))


class _Interface(object):
    pass


_uno = _module("uno")
_unohelper = _module("unohelper")
_unohelper.Base = _Base
_unohelper.ImplementationHelper = _ImplementationHelper

_module("com")
_module("com.sun", "com")
_module("com.sun.star", "com.sun")

_awt = _module("com.sun.star.awt", "com.sun.star")
_awt_type = _module("com.sun.star.awt.MessageBoxType", "com.sun.star.awt")
_awt_buttons = _module("com.sun.star.awt.MessageBoxButtons", "com.sun.star.awt")
_frame = _module("com.sun.star.frame", "com.sun.star")
_lang = _module("com.sun.star.lang", "com.sun.star")
_style = _module("com.sun.star.style", "com.sun.star")
_task = _module("com.sun.star.task", "com.sun.star")
_text = _module("com.sun.star.text", "com.sun.star")
_text_ha = _module("com.sun.star.text.HorizontalAdjust", "com.sun.star.text")
_text_wm = _module("com.sun.star.text.WritingMode2", "com.sun.star.text")
_style_pa = _module("com.sun.star.style.ParagraphAdjust", "com.sun.star.style")

_task.XJobExecutor = type("XJobExecutor", (_Interface,), {})
_frame.XDispatchProvider = type("XDispatchProvider", (_Interface,), {})
_frame.XDispatch = type("XDispatch", (_Interface,), {})
_frame.FeatureStateEvent = type("FeatureStateEvent", (object,), {})
_lang.XInitialization = type("XInitialization", (_Interface,), {})
_lang.XServiceInfo = type("XServiceInfo", (_Interface,), {})
_text_ha.LEFT = 0
_text_ha.RIGHT = 2
_text_wm.LR_TB = 0
_text_wm.RL_TB = 1
_style_pa.LEFT = 0
_style_pa.RIGHT = 1
_awt_type.MESSAGEBOX = 0
_awt_buttons.BUTTONS_OK = 0


# ---------------------------------------------------------------------------
# 2. Simulation of the Writer text model
# ---------------------------------------------------------------------------
class Anchor(object):
    """One character wide placeholder standing for a footnote anchor."""

    def __init__(self, footnote):
        self.footnote = footnote

    def render(self):
        # Writer shows the footnote number at the anchor, one character wide
        return "1"


class Position(object):
    """Live zero length position inside a FakeText."""

    def __init__(self, text, index):
        self.text = text
        self.s = index
        self.e = index
        text.ranges.append(self)


class Range(object):
    """Live text range, equivalent of com.sun.star.text.XTextRange."""

    def __init__(self, text, start, end):
        self.text = text
        self.s = start
        self.e = end
        text.ranges.append(self)

    def getString(self):
        return self.text.string_of(self.s, self.e)

    def setString(self, value):
        self.text.splice(self.s, self.e, value, owner=self)

    def getStart(self):
        return Position(self.text, self.s)

    def getEnd(self):
        return Position(self.text, self.e)

    def __repr__(self):
        return "<Range %d:%d %r>" % (self.s, self.e, self.getString())


class Cursor(Range):
    """Live text cursor, equivalent of com.sun.star.text.XTextCursor."""

    def __init__(self, text, start, end):
        Range.__init__(self, text, start, end)
        self.props = dict(text.probe_props)

    # -- XTextCursor --------------------------------------------------------
    def collapseToStart(self):
        self.e = self.s

    def collapseToEnd(self):
        self.s = self.e

    def isCollapsed(self):
        return self.s == self.e

    def goLeft(self, count, expand):
        previous = self.s
        self.s = max(0, self.s - count)
        if not expand:
            self.e = self.s
        return self.s != previous

    def goRight(self, count, expand):
        previous = self.e
        self.e = min(self.text.length(), self.e + count)
        if not expand:
            self.s = self.e
        return self.e != previous

    def gotoStart(self, expand):
        if expand:
            self.s = 0
        else:
            self.s = self.e = 0
        return True

    def gotoEnd(self, expand):
        if expand:
            self.e = self.text.length()
        else:
            self.s = self.e = self.text.length()
        return True

    def gotoRange(self, range_, expand):
        if expand:
            # bSelect = True extends the current range to include xRange
            self.s = min(self.s, range_.s)
            self.e = max(self.e, range_.e)
        else:
            self.s = self.e = range_.e
        return True

    def gotoNextParagraph(self, expand):
        index = self.text.content.find("\n", self.e)
        if index < 0:
            return False
        if expand:
            self.e = index + 1
        else:
            self.s = self.e = index + 1
        return True

    def setString(self, value):
        self.text.splice(self.s, self.e, value, owner=self)

    def getText(self):
        return self.text

    # -- format properties used by the extension ---------------------------
    def setPropertyToDefault(self, name):
        self.props.pop(name, None)

    def __getattr__(self, name):
        props = self.__dict__.get("props")
        if props and name in props:
            return props[name]
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name in ("text", "s", "e", "props"):
            object.__setattr__(self, name, value)
        else:
            self.props[name] = value


class FakeText(object):
    """Very small text model: a character list plus live anchors and ranges."""

    def __init__(self, content="", probe_props=None):
        self.chars = list(content)
        self.anchors = []
        self.ranges = []
        # Effective character properties reported when a probe cursor is
        # created (stand-in for the real formatting at that point in the body)
        self.probe_props = dict(probe_props or {})

    # -- helpers -----------------------------------------------------------
    def length(self):
        return len(self.chars)

    @property
    def content(self):
        return "".join(ch if isinstance(ch, str) else ch.render() for ch in self.chars)

    def string_of(self, start, end):
        return "".join(
            ch if isinstance(ch, str) else ch.render() for ch in self.chars[start:end]
        )

    def index_of(self, anchor):
        return self.chars.index(anchor)

    def splice(self, start, end, value, owner=None):
        inserted = list(value)
        delta = len(inserted) - (end - start)
        self.chars[start:end] = inserted
        for range_ in self.ranges:
            if range_ is owner:
                continue
            range_.s = self._shift(range_.s, start, end, delta)
            range_.e = self._shift(range_.e, start, end, delta)
        remaining = []
        for anchor in self.anchors:
            try:
                index = self.chars.index(anchor)
            except ValueError:
                continue  # this edit deleted the anchor (and its footnote)
            remaining.append(anchor)
            new_index = self._shift_anchor(index, start, end, delta)
            if new_index != index:
                self.chars.remove(anchor)
                self.chars.insert(new_index, anchor)
        self.anchors = remaining
        if owner is not None:
            owner.s = start
            owner.e = start + len(inserted)

    @staticmethod
    def _shift(value, start, end, delta):
        """Positions do not move when text is inserted exactly at them."""
        if value <= start:
            return value
        if value >= end:
            return value + delta
        return start

    @staticmethod
    def _shift_anchor(index, start, end, delta):
        """Anchors behave like ordinary characters: text inserted at the
        anchor position lands on the anchor's left side."""
        if index < start:
            return index
        return start + (delta if delta > 0 else 0)

    def insert_anchor(self, position, anchor, owner):
        self.chars.insert(position, anchor)
        self.anchors.append(anchor)
        for range_ in self.ranges:
            if range_ is owner:
                continue
            range_.s = self._shift(range_.s, position, position + 1, 1)
            range_.e = self._shift(range_.e, position, position + 1, 1)
        for other in self.anchors:
            if other is anchor:
                continue
            index = self.index_of(other)
            if index >= position:
                self.chars.remove(other)
                self.chars.insert(index + 1, other)
        if owner is not None:
            # Writer leaves the cursor behind the newly inserted anchor
            owner.s = position + 1
            owner.e = position + 1

    # -- XText -------------------------------------------------------------
    def getString(self):
        return self.content

    def createTextCursor(self):
        return Cursor(self, 0, 0)

    def createTextCursorByRange(self, range_):
        return Cursor(self, range_.s, range_.e)

    def insertString(self, cursor, value, absorb):
        position = cursor.s if cursor.s == cursor.e else cursor.e
        self.splice(position, position, value, owner=cursor)
        if not absorb:
            # bAbsorb = False leaves the cursor collapsed behind the new text
            cursor.s = cursor.e = position + len(value)

    def insertTextContent(self, cursor, content, absorb):
        if not isinstance(content, Footnote):
            raise TypeError("unsupported text content: %r" % (content,))
        content.body = self
        anchor = Anchor(content)
        content.anchor = anchor
        self.insert_anchor(cursor.e, anchor, cursor)


class Footnote(object):
    def __init__(self):
        self.body = None
        self.anchor = None
        self.footer = FakeText()

    def getAnchor(self):
        return AnchorRange(self.body, self)

    def getText(self):
        return self.footer

    def getString(self):
        return self.footer.content

    def setString(self, value):
        self.footer.chars = list(value)


class AnchorRange(object):
    """Anchor range whose bounds follow the anchor's live position."""

    def __init__(self, text, footnote):
        self.text = text
        self.footnote = footnote

    def getStart(self):
        return Position(self.text, self.text.index_of(self.footnote.anchor))

    def getEnd(self):
        return Position(self.text, self.text.index_of(self.footnote.anchor) + 1)

    def getString(self):
        return "1"


class IndexAccess(object):
    def __init__(self, items=None):
        self._items = list(items or [])

    def getCount(self):
        return len(self._items)

    def getByIndex(self, index):
        return self._items[index]


class Footnotes(object):
    """Footnote collection; always reports footnotes in document order."""

    def __init__(self, document):
        self.document = document

    def getCount(self):
        return len(self.document.footnotes_in_order())

    def getByIndex(self, index):
        return self.document.footnotes_in_order()[index]


class Bag(object):
    """Simple attribute container standing for a UNO property set."""

    def __init__(self, **kwargs):
        self.__dict__.update(kwargs)


class StyleCollection(object):
    def __init__(self, styles):
        self._styles = styles

    def hasByName(self, name):
        return name in self._styles

    def getByName(self, name):
        return self._styles[name]

    def getElementNames(self):
        return tuple(self._styles)


class StyleFamilies(object):
    def __init__(self, families):
        self._families = families

    def hasByName(self, name):
        return name in self._families

    def getByName(self, name):
        return self._families[name]


class UndoManager(object):
    def __init__(self):
        self.stack = []
        self.entered = []
        self.left = 0

    def enterUndoContext(self, title):
        self.stack.append(title)
        self.entered.append(title)

    def leaveUndoContext(self):
        self.left += 1
        if self.stack:
            self.stack.pop()


# ---------------------------------------------------------------------------
# 3. Simulation of the document and the UNO service environment
# ---------------------------------------------------------------------------
def default_styles(body_font_complex="Traditional Arabic", body_size_complex=18.0,
                   body_font_western="Times New Roman", body_size_western=12.0,
                   body_font_resolved=True):
    """Paragraph/page styles of a typical Arabic academic document."""
    standard = Bag(
        CharFontNameComplex=body_font_complex if body_font_resolved else "",
        CharHeightComplex=body_size_complex if body_font_resolved else 0.0,
        CharFontName=body_font_western if body_font_resolved else "",
        CharHeight=body_size_western if body_font_resolved else 0.0,
    )
    footnote = Bag(CharFontNameComplex="", CharHeightComplex=0.0,
                   CharFontName="", CharHeight=0.0)
    return StyleFamilies({
        "ParagraphStyles": StyleCollection({
            "Standard": standard,
            "Default Paragraph Style": standard,
            "Footnote": footnote,
            "footnote": footnote,
        }),
        "PageStyles": StyleCollection({
            "Standard": Bag(FootnoteLineAdjust=0),
            "Converted1": Bag(FootnoteLineAdjust=0),
        }),
    })


class FakeDocument(object):
    def __init__(self, content, probe_props=None, styles=None):
        self.body = FakeText(content, probe_props)
        self.footnotes = []
        self.FootnoteSettings = Bag(Prefix="", Suffix="", NumberingType=0,
                                    FootnoteCounting=1)
        self.styles = styles if styles is not None else default_styles()
        self.undo = UndoManager()
        self.locked = 0
        self.action_locks = 0
        self.view_cursor = Cursor(self.body, 0, 0)
        self.current_selection = None

    # -- the UNO methods used by the extension -----------------------------
    def supportsService(self, name):
        return name == "com.sun.star.text.TextDocument"

    def getText(self):
        return self.body

    def createSearchDescriptor(self):
        return Bag(SearchRegularExpression=False, SearchString="")

    def findAll(self, search):
        matches = []
        if getattr(search, "SearchRegularExpression", False):
            for match in re.finditer(search.SearchString, self.body.content):
                matches.append(Range(self.body, match.start(), match.end()))
        else:
            position = 0
            while True:
                index = self.body.content.find(search.SearchString, position)
                if index < 0:
                    break
                matches.append(Range(self.body, index, index + len(search.SearchString)))
                position = index + len(search.SearchString)
        return IndexAccess(matches)

    def footnotes_in_order(self):
        alive = [fn for fn in self.footnotes if fn.anchor in self.body.anchors]
        return sorted(alive, key=lambda fn: self.body.index_of(fn.anchor))

    def getFootnotes(self):
        return Footnotes(self)

    def createInstance(self, name):
        if name != "com.sun.star.text.Footnote":
            raise Exception("unsupported text content: " + name)
        footnote = Footnote()
        self.footnotes.append(footnote)
        return footnote

    def getStyleFamilies(self):
        return self.styles

    def getUndoManager(self):
        return self.undo

    def lockControllers(self):
        self.locked += 1

    def unlockControllers(self):
        self.locked -= 1

    def addActionLock(self):
        self.action_locks += 1

    def removeActionLock(self):
        self.action_locks -= 1

    def getCurrentController(self):
        return Bag(getViewCursor=lambda: self.view_cursor)

    def getCurrentSelection(self):
        return self.current_selection

    # -- helpers for the assertions ----------------------------------------
    def body_text(self):
        return self.body.content

    def footnote_texts(self):
        return [fn.getString() for fn in self.footnotes_in_order()]


class FakeDesktop(object):
    def __init__(self, document):
        self.document = document

    def getCurrentComponent(self):
        return self.document


class FakeServiceManager(object):
    def __init__(self, document):
        self.document = document

    def createInstanceWithContext(self, name, context):
        if name == "com.sun.star.frame.Desktop":
            return FakeDesktop(self.document)
        raise Exception("unsupported service: " + name)


class FakeContext(object):
    def __init__(self, document):
        self.SM = FakeServiceManager(document)

    def getServiceManager(self):
        return self.SM


# ---------------------------------------------------------------------------
# 4. Load the extension under test
# ---------------------------------------------------------------------------
MESSAGES = []


def load_extension():
    spec = importlib.util.spec_from_file_location("footnoteconverter_under_test",
                                                  EXTENSION)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    # Show messages in the console instead of a real message box
    module.FootnoteService.show_message = lambda self, title, message: MESSAGES.append(
        (title, message)
    )
    return module


MODULE = load_extension()


def make_service(document):
    return MODULE.FootnoteService(FakeContext(document))


# ---------------------------------------------------------------------------
# 5. Tests
# ---------------------------------------------------------------------------
CHECKS = []

SAMPLE = "قال «جلّ جلاله» ثم <الكشاف> وقال ((صلى الله عليه وسلم)) و[البخاري] ."
SAMPLE_CONVERTED = "قال «جلّ جلاله» ثم(1) وقال ((صلى الله عليه وسلم)) و[البخاري] ."


def check(name, condition, detail=""):
    CHECKS.append((name, bool(condition), detail))
    print("[%s] %s%s" % ("PASS" if condition else "FAIL", name,
                         "" if condition else "  -> " + detail))


def _search(document):
    search = document.createSearchDescriptor()
    search.SearchRegularExpression = True
    search.SearchString = MODULE.COMBINED_REGEX
    return document.findAll(search)


def test_marker_protection():
    """Only < > is a footnote marker; « », (( )) and [ ] must stay untouched."""
    document = FakeDocument(SAMPLE)
    matches = _search(document)
    check("only < > is matched as a footnote marker",
          matches.getCount() == 1, "matches=%d" % matches.getCount())
    if matches.getCount():
        check("the matched span is exactly the < > marker",
              matches.getByIndex(0).getString() == "<الكشاف>",
              repr(matches.getByIndex(0).getString()))


def test_conversion():
    """The marker becomes a real footnote with a glued (1) reference."""
    document = FakeDocument(SAMPLE)
    service = make_service(document)
    service.convert_footnotes()

    check("one footnote was created", document.getFootnotes().getCount() == 1,
          "count=%d" % document.getFootnotes().getCount())
    check("the footnote text moved below the page",
          document.footnote_texts() == ["الكشاف"], repr(document.footnote_texts()))
    check("the < > marker was replaced by (1) glued to the previous word",
          document.body_text() == SAMPLE_CONVERTED, repr(document.body_text()))
    check("the document was unlocked again after the operation",
          document.locked == 0 and document.action_locks == 0,
          "controllers=%d action_locks=%d" % (document.locked, document.action_locks))
    check("the whole operation is undoable in one step",
          document.undo.entered == ["Convert Footnotes | تحويل الحواشي"]
          and document.undo.left == 1, repr(document.undo.entered))
    check("the success message reports one converted footnote",
          bool(MESSAGES) and "تم تحويل (1)" in MESSAGES[-1][1],
          repr(MESSAGES[-1:]))


def test_second_run_is_idempotent():
    """Running the converter twice must not duplicate parentheses or footnotes."""
    document = FakeDocument(SAMPLE)
    service = make_service(document)
    service.convert_footnotes()
    first_pass = document.body_text()
    service.convert_footnotes()

    check("a second run creates no extra footnote",
          document.getFootnotes().getCount() == 1,
          "count=%d" % document.getFootnotes().getCount())
    check("a second run does not duplicate the parentheses",
          document.body_text() == first_pass, repr(document.body_text()))
    check("a second run reports that no marker was found",
          "لم يتم العثور" in MESSAGES[-1][1], repr(MESSAGES[-1:]))


def test_already_parenthesized_marker():
    """A marker that is already wrapped by ( ) reuses that pair."""
    document = FakeDocument("كلمة (<حاشية>) أخرى")
    service = make_service(document)
    service.convert_footnotes()

    check("the existing parentheses are reused instead of duplicated",
          document.body_text() == "كلمة (1) أخرى", repr(document.body_text()))
    check("the footnote itself was still created",
          document.footnote_texts() == ["حاشية"], repr(document.footnote_texts()))


def test_empty_marker_cleanup():
    """<> must vanish from the body without creating an empty footnote."""
    document = FakeDocument("قال تعالى <> ثم أكمل")
    service = make_service(document)
    service.convert_footnotes()

    check("an empty marker creates no footnote",
          document.getFootnotes().getCount() == 0,
          "count=%d" % document.getFootnotes().getCount())
    check("the empty marker and the space before it are cleaned up",
          document.body_text() == "قال تعالى ثم أكمل", repr(document.body_text()))


def test_punctuation_gluing():
    """A punctuation mark after the marker moves right behind the reference."""
    document = FakeDocument("كلمة <حاشية> .")
    service = make_service(document)
    service.convert_footnotes()

    check("the punctuation is glued directly after the reference",
          document.body_text() == "كلمة(1).", repr(document.body_text()))


def test_reverse_conversion():
    """Footnotes can be extracted back into < > markers."""
    document = FakeDocument(SAMPLE)
    service = make_service(document)
    service.convert_footnotes()
    service.reverse_footnotes()

    check("the footnote text returns to the body as a < > marker",
          document.body_text() == SAMPLE, repr(document.body_text()))
    check("the footnote area is empty afterwards",
          document.getFootnotes().getCount() == 0,
          "count=%d" % document.getFootnotes().getCount())
    check("the reverse operation is undoable in one step",
          document.undo.entered[-1] == "Reverse Footnotes | استخراج الحواشي",
          repr(document.undo.entered))


def test_round_trip_is_stable():
    """convert -> reverse -> convert must reproduce the first result."""
    document = FakeDocument(SAMPLE)
    service = make_service(document)
    service.convert_footnotes()
    first_pass = document.body_text()
    service.reverse_footnotes()
    service.convert_footnotes()

    check("the convert/reverse round trip is stable",
          document.body_text() == first_pass and
          document.footnote_texts() == ["الكشاف"],
          repr(document.body_text()))


def test_arabic_document_detection():
    service = make_service(FakeDocument(""))
    latin_cover = "Master thesis submitted to the department of Islamic studies. " * 60

    check("Arabic text is recognised",
          service.is_arabic_text("بسم الله الرحمن الرحيم"))
    check("Latin text is not recognised as Arabic",
          not service.is_arabic_text("Footnote Converter"))
    check("an empty string is not recognised as Arabic",
          not service.is_arabic_text(""))
    check("a purely Arabic document is treated as Arabic",
          service.is_arabic_document(FakeDocument("بسم الله <حاشية>")))
    check("Arabic after a long Latin cover page is still detected",
          service.is_arabic_document(
              FakeDocument(latin_cover + " وقال الإمام النووي رحمه الله")))
    check("a purely Latin document is not treated as Arabic",
          not service.is_arabic_document(FakeDocument(latin_cover)))


def test_unify_font_size_from_style():
    """The footnote size follows the body style: body size - 2 pt."""
    document = FakeDocument("كلمة <حاشية> أخرى")
    service = make_service(document)
    service.convert_footnotes()
    del MESSAGES[:]
    service.unify_footnotes_font()

    footnote_style = document.styles.getByName("ParagraphStyles").getByName("Footnote")
    check("the footnote style is two points smaller than the body text",
          footnote_style.CharHeightComplex == 16.0,
          "size=%r" % (footnote_style.CharHeightComplex,))
    check("the footnote style inherits the body font",
          footnote_style.CharFontNameComplex == "Traditional Arabic",
          repr(footnote_style.CharFontNameComplex))
    check("the Footnote style is switched to right to left for Arabic",
          footnote_style.WritingMode == 1 and footnote_style.ParaAdjust == 1,
          "writing mode=%r adjust=%r" % (footnote_style.WritingMode,
                                         footnote_style.ParaAdjust))
    check("the message reports the resulting size",
          bool(MESSAGES) and "16" in MESSAGES[-1][1], repr(MESSAGES[-1:]))
    check("the font unification is undoable in one step",
          document.undo.entered[-1] == "Unify Footnotes Font & Size | توحيد خط الحواشي"
          and document.undo.left == len(document.undo.entered),
          repr(document.undo.entered))
    check("the document is unlocked after the unification",
          document.locked == 0 and document.action_locks == 0,
          "controllers=%d action_locks=%d" % (document.locked, document.action_locks))


def test_unify_font_size_from_body_probe():
    """When the styles do not expose a size the real body text is measured."""
    styles = default_styles(body_font_resolved=False)
    document = FakeDocument(
        "كلمة <حاشية> أخرى",
        probe_props={"CharFontNameComplex": "Simplified Arabic",
                     "CharHeightComplex": 14.0,
                     "CharFontName": "Arial",
                     "CharHeight": 11.0},
        styles=styles,
    )
    service = make_service(document)
    service.convert_footnotes()
    del MESSAGES[:]
    service.unify_footnotes_font()

    footnote_style = styles.getByName("ParagraphStyles").getByName("Footnote")
    check("the measured body size is used instead of the 18 pt default",
          footnote_style.CharHeightComplex == 12.0,
          "size=%r" % (footnote_style.CharHeightComplex,))
    check("the measured body font is applied to the footnotes",
          footnote_style.CharFontNameComplex == "Simplified Arabic",
          repr(footnote_style.CharFontNameComplex))


def test_numbering_scope():
    """Restart per page vs continuous numbering."""
    document = FakeDocument("كلمة <حاشية>")
    service = make_service(document)

    service.set_numbering_scope(MODULE.PER_PAGE)
    check("per page numbering sets FootnoteCounting to 0",
          document.FootnoteSettings.FootnoteCounting == 0,
          repr(document.FootnoteSettings.FootnoteCounting))

    service.set_numbering_scope(MODULE.PER_DOCUMENT)
    check("continuous numbering sets FootnoteCounting to 2",
          document.FootnoteSettings.FootnoteCounting == 2,
          repr(document.FootnoteSettings.FootnoteCounting))


def test_format_numbering_wraps_existing_footnotes():
    """Format & Align wraps already existing footnotes and is undoable."""
    document = FakeDocument("كلمة  أخرى")
    footnote = document.createInstance("com.sun.star.text.Footnote")
    position = document.body.createTextCursor()
    position.gotoStart(False)
    position.goRight(4, False)
    document.body.insertTextContent(position, footnote, False)
    footnote.setString("حاشية")

    service = make_service(document)
    del MESSAGES[:]
    service.format_numbering()

    check("an existing footnote receives its ( ) reference",
          document.body_text() == "كلمة(1)  أخرى", repr(document.body_text()))
    check("the footnote numbering format is set to (1)",
          document.FootnoteSettings.Prefix == "("
          and document.FootnoteSettings.Suffix == ") "
          and document.FootnoteSettings.NumberingType == 4,
          repr((document.FootnoteSettings.Prefix, document.FootnoteSettings.Suffix,
                document.FootnoteSettings.NumberingType)))
    check("the separator line is moved to the right for Arabic",
          document.styles.getByName("PageStyles").getByName("Standard").FootnoteLineAdjust == 2,
          repr(document.styles.getByName("PageStyles").getByName("Standard").FootnoteLineAdjust))
    check("the message reports the number of fixed footnotes",
          bool(MESSAGES) and "إحاطة (1)" in MESSAGES[-1][1], repr(MESSAGES[-1:]))
    check("format & align runs in a single undo context",
          document.undo.entered[-1] == "Format & Align | ضبط التنسيق والمحاذاة"
          and document.undo.left == len(document.undo.entered),
          repr(document.undo.entered))
    check("the document is unlocked after format & align",
          document.locked == 0 and document.action_locks == 0,
          "controllers=%d action_locks=%d" % (document.locked, document.action_locks))

    del MESSAGES[:]
    service.format_numbering()
    check("a second format run changes nothing and fixes nothing",
          document.body_text() == "كلمة(1)  أخرى" and "إحاطة (0)" in MESSAGES[-1][1],
          repr((document.body_text(), MESSAGES[-1:])))


def test_wrap_selection():
    """F2 wraps the selection in < > and strips outer parentheses."""
    document = FakeDocument("كلمة ثم نص")
    document.current_selection = IndexAccess([Range(document.body, 0, 4)])
    document.view_cursor = Cursor(document.body, 9, 9)
    service = make_service(document)
    service.wrap_selection_with_markers()
    check("the selected text is wrapped in < >",
          document.body_text() == "<كلمة> ثم نص", repr(document.body_text()))

    second = FakeDocument("(نص)")
    second.current_selection = IndexAccess([Range(second.body, 0, 5)])
    second.view_cursor = Cursor(second.body, 5, 5)
    make_service(second).wrap_selection_with_markers()
    check("outer parentheses are replaced by the < > marker",
          second.body_text() == "<نص>", repr(second.body_text()))

    third = FakeDocument("")
    third.current_selection = IndexAccess([])
    cursor = Cursor(third.body, 0, 0)
    third.view_cursor = cursor
    make_service(third).wrap_selection_with_markers()
    check("with no selection an empty <> pair is inserted",
          third.body_text() == "<>" and cursor.e == 1,
          repr((third.body_text(), cursor.s, cursor.e)))


# ---------------------------------------------------------------------------
# 6. Runner
# ---------------------------------------------------------------------------
TESTS = [
    ("marker protection", test_marker_protection),
    ("conversion", test_conversion),
    ("second run is idempotent", test_second_run_is_idempotent),
    ("already parenthesized marker", test_already_parenthesized_marker),
    ("empty marker cleanup", test_empty_marker_cleanup),
    ("punctuation gluing", test_punctuation_gluing),
    ("reverse conversion", test_reverse_conversion),
    ("round trip is stable", test_round_trip_is_stable),
    ("arabic document detection", test_arabic_document_detection),
    ("unify font size from style", test_unify_font_size_from_style),
    ("unify font size from body probe", test_unify_font_size_from_body_probe),
    ("numbering scope", test_numbering_scope),
    ("format numbering", test_format_numbering_wraps_existing_footnotes),
    ("wrap selection", test_wrap_selection),
]


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    print("Footnote Converter Pro - logic self-test")
    print("=" * 70)
    for name, function in TESTS:
        del MESSAGES[:]
        try:
            function()
        except Exception as error:  # one broken test must not hide the others
            check(name + ": no unexpected exception", False, repr(error))

    failed = [item for item in CHECKS if not item[1]]
    print("=" * 70)
    print("checks: %d   failed: %d" % (len(CHECKS), len(failed)))
    for name, _, detail in failed:
        print("FAILED: %s -> %s" % (name, detail))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())