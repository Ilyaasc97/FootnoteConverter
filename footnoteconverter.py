# -*- coding: utf-8 -*-
"""
Footnote Converter Pro Extension for LibreOffice Writer (Version 2.3.1)
Bilingual Edition (Arabic & English) | النسخة ثنائية اللغة (عربي وإنجليزي)

Exclusive marker / علامة الحصر: ONLY < ... > is converted.
Quotes « », hadith brackets (( )) and square brackets [ ] are ALWAYS protected.
تُحوَّل العلامات < > فقط، ولا تُمس أقواس الاقتباس « » ولا أقواس الحديث (( )) ولا الأقواس [ ].

Features / المميزات:
1. Converts < ... > markers into native footnotes with (1) numbering in the text
   body and (1) numbering in the footnote footer.
2. Smart space and punctuation cleanup (attaches footnote directly to word without gap).
3. Numbering scope controls (Restart per page vs Continuous per document).
4. Auto-alignment of footnote separator line and footnote paragraphs to RTL/Right for Arabic.
5. Reverse conversion: Extract existing footnotes back into text < > markers.
6. Font unifier: footnote font = body font, and footnote size = body size - 2 pt.
7. Full Undo (Ctrl + Z) support for every modifying operation.
"""

import uno
import unohelper
from com.sun.star.task import XJobExecutor
from com.sun.star.frame import XDispatchProvider, XDispatch
from com.sun.star.lang import XInitialization, XServiceInfo

PROTOCOL_HANDLER_NAME = "com.github.ilyaasc97.footnoteconverter.ProtocolHandler"
JOB_EXECUTOR_NAME = "com.github.ilyaasc97.footnoteconverter.FootnoteConverter"

try:
    from com.sun.star.frame import FeatureStateEvent
except ImportError:
    FeatureStateEvent = None

try:
    from com.sun.star.text.HorizontalAdjust import LEFT, RIGHT
except ImportError:
    LEFT = 0
    RIGHT = 2

try:
    from com.sun.star.text.WritingMode2 import LR_TB, RL_TB
except ImportError:
    LR_TB = 0
    RL_TB = 1

try:
    from com.sun.star.style.ParagraphAdjust import LEFT as PARA_LEFT, RIGHT as PARA_RIGHT
except ImportError:
    PARA_LEFT = 0
    PARA_RIGHT = 1

PER_PAGE = 0
PER_CHAPTER = 1
PER_DOCUMENT = 2

COMBINED_REGEX = r"<([^<>]*)>"


class FootnoteService:
    def __init__(self, ctx):
        self.ctx = ctx
        self.smgr = ctx.getServiceManager() if hasattr(ctx, "getServiceManager") else ctx.ServiceManager

    def get_current_doc(self):
        desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        doc = desktop.getCurrentComponent()
        if not doc and hasattr(desktop, "getComponents"):
            try:
                components = desktop.getComponents().createEnumeration()
                while components.hasMoreElements():
                    comp = components.nextElement()
                    if hasattr(comp, "supportsService") and comp.supportsService("com.sun.star.text.TextDocument"):
                        doc = comp
                        break
            except Exception:
                pass
        return doc

    def show_message(self, title, message):
        try:
            toolkit = self.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
            from com.sun.star.awt.MessageBoxType import MESSAGEBOX
            from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
            box = toolkit.createMessageBox(None, MESSAGEBOX, BUTTONS_OK, title, message)
            box.execute()
        except Exception as e:
            print(f"[FootnoteConverter] show_message error: {e}")

    def is_arabic_text(self, text):
        if not text:
            return False
        for ch in text:
            if '\u0600' <= ch <= '\u06FF' or '\u0750' <= ch <= '\u077F' or '\u08A0' <= ch <= '\u08FF':
                return True
        return False

    def is_arabic_document(self, doc):
        try:
            full_text = doc.getText().getString()
            # Fast path: Arabic in the opening sample (typical Arabic documents)
            if self.is_arabic_text(full_text[:2500]):
                return True
            fns = doc.getFootnotes()
            if fns and fns.getCount() > 0:
                for i in range(min(5, fns.getCount())):
                    if self.is_arabic_text(fns.getByIndex(i).getString()):
                        return True
            # Wide scan: Arabic anywhere in the body. An Arabic thesis often
            # opens with a Latin cover page and must still be handled as Arabic.
            if self.is_arabic_text(full_text):
                return True
            # Latin fallback: clearly a non-Arabic document
            has_latin = any(('a' <= ch <= 'z') or ('A' <= ch <= 'Z') for ch in full_text[:2500])
            if has_latin:
                return False
            cur = doc.getText().createTextCursor()
            if hasattr(cur, "WritingMode"):
                return cur.WritingMode == RL_TB
        except Exception:
            pass
        return False

    def adjust_footnote_separator(self, doc, is_rtl=True):
        adjust_val = RIGHT if is_rtl else LEFT
        try:
            style_families = doc.getStyleFamilies()
            if style_families.hasByName("PageStyles"):
                page_styles = style_families.getByName("PageStyles")
                for name in page_styles.getElementNames():
                    ps = page_styles.getByName(name)
                    if hasattr(ps, "FootnoteLineAdjust"):
                        try:
                            ps.FootnoteLineAdjust = adjust_val
                        except Exception:
                            pass
        except Exception as e:
            print(f"[FootnoteConverter] adjust_footnote_separator error: {e}")

    def adjust_footnote_paragraphs(self, doc, is_rtl=True):
        wm = RL_TB if is_rtl else LR_TB
        pa = PARA_RIGHT if is_rtl else PARA_LEFT

        try:
            style_families = doc.getStyleFamilies()
            if style_families.hasByName("ParagraphStyles"):
                para_styles = style_families.getByName("ParagraphStyles")
                for style_name in ["Footnote", "footnote"]:
                    if para_styles.hasByName(style_name):
                        ps = para_styles.getByName(style_name)
                        try:
                            ps.WritingMode = wm
                            ps.ParaAdjust = pa
                        except Exception:
                            pass
        except Exception as e:
            print(f"[FootnoteConverter] adjust Footnote style error: {e}")

        try:
            fns = doc.getFootnotes()
            if fns:
                for i in range(fns.getCount()):
                    fn = fns.getByIndex(i)
                    fn_cur = fn.getText().createTextCursor()
                    fn_cur.gotoStart(False)
                    fn_cur.gotoEnd(True)
                    try:
                        fn_cur.WritingMode = wm
                        fn_cur.ParaAdjust = pa
                    except Exception:
                        pass
        except Exception as e:
            print(f"[FootnoteConverter] adjust individual footnotes error: {e}")

    def fix_existing_footnote_brackets(self, doc):
        text = doc.getText()
        fns = doc.getFootnotes()
        count = fns.getCount() if fns else 0
        fixed = 0

        for i in range(count - 1, -1, -1):
            fn = fns.getByIndex(i)
            anchor = fn.getAnchor()
            if not anchor:
                continue

            patched = False

            cur_before = text.createTextCursorByRange(anchor.getStart())
            has_before = cur_before.goLeft(1, True)
            char_before = cur_before.getString() if has_before else ""

            cur_after = text.createTextCursorByRange(anchor.getEnd())
            has_after = cur_after.goRight(1, True)
            char_after = cur_after.getString() if has_after else ""

            if char_before != "(":
                c = text.createTextCursorByRange(anchor.getStart())
                try:
                    c.CharStyleName = "Footnote anchor"
                except Exception:
                    pass
                text.insertString(c, "(", False)
                try:
                    c.setPropertyToDefault("CharStyleName")
                except Exception:
                    pass
                patched = True

            if char_after != ")":
                c = text.createTextCursorByRange(anchor.getEnd())
                try:
                    c.CharStyleName = "Footnote anchor"
                except Exception:
                    pass
                text.insertString(c, ")", False)
                try:
                    c.setPropertyToDefault("CharStyleName")
                except Exception:
                    pass
                patched = True

            # Count fixed footnotes (not the single inserted bracket characters)
            if patched:
                fixed += 1

        return fixed

    def extract_inner_content(self, raw_text):
        t = raw_text.strip()
        if t.startswith("<") and t.endswith(">"):
            return t[1:-1].strip()
        return t

    def wrap_selection_with_markers(self):
        """Wraps currently selected text in < > (stripping any outer parentheses).
        If nothing is selected, inserts <> and places cursor inside."""
        doc = self.get_current_doc()
        if not doc or not hasattr(doc, "supportsService") or not doc.supportsService("com.sun.star.text.TextDocument"):
            return

        controller = doc.getCurrentController()
        if not controller:
            return

        selection = doc.getCurrentSelection()
        view_cursor = controller.getViewCursor()

        undo_manager = None
        try:
            if hasattr(doc, "getUndoManager"):
                undo_manager = doc.getUndoManager()
                if undo_manager:
                    undo_manager.enterUndoContext("Wrap with < > | إحاطة النص")
        except Exception:
            undo_manager = None

        try:
            has_selection = False
            if selection and hasattr(selection, "getCount"):
                for idx in range(selection.getCount()):
                    range_obj = selection.getByIndex(idx)
                    raw = range_obj.getString()
                    if raw and len(raw.strip()) > 0:
                        has_selection = True
                        trimmed = raw.strip()
                        if trimmed.startswith("(") and trimmed.endswith(")"):
                            inner = trimmed[1:-1].strip()
                        elif trimmed.startswith("<") and trimmed.endswith(">"):
                            inner = trimmed[1:-1].strip()
                        else:
                            inner = trimmed

                        l_space = raw[:len(raw) - len(raw.lstrip())]
                        r_space = raw[len(raw.rstrip()):]

                        range_obj.setString(f"{l_space}<{inner}>{r_space}")

            if not has_selection and view_cursor:
                text = view_cursor.getText()
                text.insertString(view_cursor, "<>", False)
                view_cursor.goLeft(1, False)

        except Exception as e:
            print(f"[FootnoteConverter] wrap_selection error: {e}")
        finally:
            if undo_manager:
                try:
                    undo_manager.leaveUndoContext()
                except Exception:
                    pass

    def convert_footnotes(self):
        doc = self.get_current_doc()
        if not doc or not hasattr(doc, "supportsService") or not doc.supportsService("com.sun.star.text.TextDocument"):
            self.show_message(
                "Warning | تنبيه",
                "[العربية] يرجى فتح مستند نصي (LibreOffice Writer) أولاً.\n"
                "[English] Please open a text document (LibreOffice Writer) first."
            )
            return

        is_rtl = self.is_arabic_document(doc)

        # Set FootnoteSettings BEFORE entering Undo context
        # to ensure global document properties do not wipe the text Undo stack
        try:
            fn_settings = doc.FootnoteSettings
            if fn_settings.Prefix != "(" or fn_settings.Suffix != ") ":
                fn_settings.Prefix = "("
                fn_settings.Suffix = ") "
                fn_settings.NumberingType = 4
        except Exception as e:
            print(f"[FootnoteConverter] FootnoteSettings error: {e}")

        undo_manager = None
        try:
            if hasattr(doc, "getUndoManager"):
                undo_manager = doc.getUndoManager()
                if undo_manager:
                    undo_manager.enterUndoContext("Convert Footnotes | تحويل الحواشي")
        except Exception:
            undo_manager = None

        locked_controllers = False
        action_locked = False
        try:
            if hasattr(doc, "lockControllers"):
                doc.lockControllers()
                locked_controllers = True
            if hasattr(doc, "addActionLock"):
                doc.addActionLock()
                action_locked = True

            text = doc.getText()
            search = doc.createSearchDescriptor()
            search.SearchRegularExpression = True
            search.SearchString = COMBINED_REGEX
            matches = doc.findAll(search)

            converted = 0
            if matches and matches.getCount() > 0:
                total_matches = matches.getCount()
                for i in range(total_matches - 1, -1, -1):
                    m = matches.getByIndex(i)
                    raw_text = m.getString().strip()
                    content = self.extract_inner_content(raw_text)
                    if not content or len(content.strip()) == 0:
                        # Empty marker <>: clean it from body text without creating a
                        # footnote (also drop the single space that used to precede it)
                        cur_pre_empty = text.createTextCursorByRange(m.getStart())
                        if cur_pre_empty.goLeft(1, True) and cur_pre_empty.getString() in (" ", "\u00A0"):
                            cur_pre_empty.setString("")
                        m.setString("")
                        continue

                    cur_pre = text.createTextCursorByRange(m.getStart())
                    if cur_pre.goLeft(1, True) and cur_pre.getString() in (" ", "\u00A0"):
                        cur_pre.setString("")

                    cur_post = text.createTextCursorByRange(m.getEnd())
                    if cur_post.goRight(2, True):
                        post_str = cur_post.getString()
                        if len(post_str) == 2 and post_str[0] in (" ", "\u00A0") and post_str[1] in (".", "،", ":", "؛", "!", "؟", ","):
                            cur_post.setString(post_str[1])

                    m_start = m.getStart()
                    m_end = m.getEnd()

                    cur_chk_before = text.createTextCursorByRange(m_start)
                    had_pre_open = cur_chk_before.goLeft(1, True) and cur_chk_before.getString() == "("

                    cur_chk_after = text.createTextCursorByRange(m_end)
                    had_post_close = cur_chk_after.goRight(1, True) and cur_chk_after.getString() == ")"

                    if had_pre_open and had_post_close:
                        cur_absorb = text.createTextCursorByRange(m_start)
                        cur_absorb.goLeft(1, False)
                        cur_absorb.gotoRange(m_end, True)
                        cur_absorb.goRight(1, True)
                        cur_absorb.setString("")
                        insert_pos = cur_absorb.getStart()
                    else:
                        m.setString("")
                        insert_pos = m.getStart()

                    cur_insert = text.createTextCursorByRange(insert_pos)
                    try:
                        cur_insert.CharStyleName = "Footnote anchor"
                    except Exception:
                        pass

                    text.insertString(cur_insert, "(", False)
                    cur_insert.collapseToEnd()

                    footnote = doc.createInstance("com.sun.star.text.Footnote")
                    text.insertTextContent(cur_insert, footnote, False)
                    cur_insert.collapseToEnd()

                    try:
                        cur_insert.CharStyleName = "Footnote anchor"
                    except Exception:
                        pass

                    text.insertString(cur_insert, ")", False)
                    cur_insert.collapseToEnd()

                    try:
                        cur_insert.setPropertyToDefault("CharStyleName")
                    except Exception:
                        pass

                    footnote.setString(content)

                    if is_rtl:
                        try:
                            fn_c = footnote.getText().createTextCursor()
                            fn_c.gotoStart(False)
                            fn_c.gotoEnd(True)
                            fn_c.WritingMode = RL_TB
                            fn_c.ParaAdjust = PARA_RIGHT
                        except Exception:
                            pass

                    converted += 1

            fixed_existing = self.fix_existing_footnote_brackets(doc)

            # Unlock document and seal the undo context BEFORE the modal message box
            if action_locked and hasattr(doc, "removeActionLock"):
                try:
                    doc.removeActionLock()
                    action_locked = False
                except Exception:
                    pass
            if locked_controllers and hasattr(doc, "unlockControllers"):
                try:
                    doc.unlockControllers()
                    locked_controllers = False
                except Exception:
                    pass
            if undo_manager:
                try:
                    undo_manager.leaveUndoContext()
                    undo_manager = None
                except Exception as ex:
                    print(f"[FootnoteConverter] leaveUndoContext error: {ex}")

            if converted > 0 or fixed_existing > 0:
                msg = (
                    f"تمت العملية بنجاح! | Success!\n"
                    f"-----------------------------\n"
                    f"[ العربية ]\n"
                    f"• تم تحويل ({converted}) حاشية سفلية جديدة.\n"
                    f"• ترقيم متن النص بين قوسين (1) ملتصقاً بالكلمة.\n"
                    f"• ضبط محاذاة الحواشي وخط الفاصل حسب اتجاه اللغة.\n"
                    f"• يمكنك التراجع في أي وقت عبر (Ctrl + Z).\n\n"
                    f"[ English ]\n"
                    f"• Successfully converted ({converted}) new footnote(s).\n"
                    f"• Footnote reference in text enclosed in (1).\n"
                    f"• Footnotes and separator line aligned to the document language.\n"
                    f"• You can undo at any time using (Ctrl + Z)."
                )
            else:
                msg = (
                    "[ العربية ] لم يتم العثور على أية نصوص داخل علامات الحصر في المستند.\n\n"
                    "[ English ] No text found between footnote markers in the document."
                )

            self.show_message("Footnote Converter | محول الحواشي", msg)

        except Exception as e:
            self.show_message("Error | خطأ", f"Error | خطأ:\n{str(e)}")
        finally:
            if action_locked and hasattr(doc, "removeActionLock"):
                try:
                    doc.removeActionLock()
                except Exception:
                    pass
            if locked_controllers and hasattr(doc, "unlockControllers"):
                try:
                    doc.unlockControllers()
                except Exception:
                    pass
            if undo_manager:
                try:
                    undo_manager.leaveUndoContext()
                except Exception:
                    pass

    def set_numbering_scope(self, scope_type):
        doc = self.get_current_doc()
        if not doc or not hasattr(doc, "supportsService") or not doc.supportsService("com.sun.star.text.TextDocument"):
            self.show_message(
                "Warning | تنبيه",
                "[العربية] يرجى فتح مستند نصي (LibreOffice Writer) أولاً.\n"
                "[English] Please open a text document (LibreOffice Writer) first."
            )
            return

        try:
            fn_settings = doc.FootnoteSettings
            fn_settings.FootnoteCounting = scope_type
            if scope_type == PER_PAGE:
                msg = (
                    "[ العربية ]\n"
                    "تم بنجاح ضبط ترقيم الحواشي ليكون: (مستقلاً يبدأ من 1 في كل صفحة جديدة).\n\n"
                    "[ English ]\n"
                    "Footnote numbering set to: Restart on each page (starts from 1)."
                )
            else:
                msg = (
                    "[ العربية ]\n"
                    "تم بنجاح ضبط ترقيم الحواشي ليكون: (ترقيماً متسلسلاً مستمراً لكامل المستند).\n\n"
                    "[ English ]\n"
                    "Footnote numbering set to: Continuous (throughout the whole document)."
                )
            self.show_message("Numbering Scope | نطاق الترقيم", msg)
        except Exception as e:
            self.show_message("Error | خطأ", f"Error | خطأ:\n{str(e)}")

    def reverse_footnotes(self):
        doc = self.get_current_doc()
        if not doc or not hasattr(doc, "supportsService") or not doc.supportsService("com.sun.star.text.TextDocument"):
            self.show_message(
                "Warning | تنبيه",
                "[العربية] يرجى فتح مستند نصي (LibreOffice Writer) أولاً.\n"
                "[English] Please open a text document (LibreOffice Writer) first."
            )
            return

        fns = doc.getFootnotes()
        count = fns.getCount() if fns else 0
        if count == 0:
            self.show_message(
                "Reverse Footnotes | التحويل العكسي",
                "[ العربية ] لا توجد أية حواشي سفلية في المستند لاستخراجها.\n"
                "[ English ] No footnotes found in the document to extract."
            )
            return

        undo_manager = None
        try:
            if hasattr(doc, "getUndoManager"):
                undo_manager = doc.getUndoManager()
                if undo_manager:
                    undo_manager.enterUndoContext("Reverse Footnotes | استخراج الحواشي")
        except Exception:
            undo_manager = None

        text = doc.getText()
        locked = False
        action_locked = False
        try:
            if hasattr(doc, "lockControllers"):
                doc.lockControllers()
                locked = True
            if hasattr(doc, "addActionLock"):
                doc.addActionLock()
                action_locked = True

            extracted = 0
            for i in range(count - 1, -1, -1):
                f = fns.getByIndex(i)
                content = f.getString().strip()
                anchor = f.getAnchor()
                if not anchor:
                    continue

                cur_b = text.createTextCursorByRange(anchor.getStart())
                has_b = cur_b.goLeft(1, True) and cur_b.getString() == "("

                cur_a = text.createTextCursorByRange(anchor.getEnd())
                has_a = cur_a.goRight(1, True) and cur_a.getString() == ")"

                if has_b and has_a:
                    cur_rep = text.createTextCursorByRange(anchor.getStart())
                    cur_rep.goLeft(1, False)
                    cur_rep.gotoRange(anchor.getEnd(), True)
                    cur_rep.goRight(1, True)
                else:
                    cur_rep = text.createTextCursorByRange(anchor.getStart())
                    cur_rep.gotoRange(anchor.getEnd(), True)

                try:
                    cur_rep.setPropertyToDefault("CharStyleName")
                except Exception:
                    pass

                cur_rep.setString(f" <{content}>")
                extracted += 1

            # Unlock document and seal the undo context BEFORE the modal message box
            if action_locked and hasattr(doc, "removeActionLock"):
                try:
                    doc.removeActionLock()
                    action_locked = False
                except Exception:
                    pass
            if locked and hasattr(doc, "unlockControllers"):
                try:
                    doc.unlockControllers()
                    locked = False
                except Exception:
                    pass
            if undo_manager:
                try:
                    undo_manager.leaveUndoContext()
                    undo_manager = None
                except Exception:
                    pass

            self.show_message(
                "Reverse Footnotes | التحويل العكسي",
                f"تمت العملية بنجاح! | Success!\n"
                f"-----------------------------\n"
                f"[ العربية ]\n"
                f"• تم بنجاح استخراج ({extracted}) حاشية سفلية وإعادتها لنصوص داخل المتن بين < >.\n"
                f"• تم تفريغ منطقة الحواشي لتسهيل نسخ النص ونشره.\n"
                f"• يمكنك التراجع في أي وقت عبر (Ctrl + Z).\n\n"
                f"[ English ]\n"
                f"• Successfully extracted ({extracted}) footnote(s) back into < > text markers.\n"
                f"• Footnote footer area has been cleared.\n"
                f"• You can undo at any time using (Ctrl + Z)."
            )

        except Exception as e:
            self.show_message("Error | خطأ", f"Error | خطأ:\n{str(e)}")
        finally:
            if action_locked and hasattr(doc, "removeActionLock"):
                try:
                    doc.removeActionLock()
                except Exception:
                    pass
            if locked and hasattr(doc, "unlockControllers"):
                try:
                    doc.unlockControllers()
                except Exception:
                    pass
            if undo_manager:
                try:
                    undo_manager.leaveUndoContext()
                except Exception:
                    pass

    def format_numbering(self):
        doc = self.get_current_doc()
        if not doc or not hasattr(doc, "supportsService") or not doc.supportsService("com.sun.star.text.TextDocument"):
            self.show_message(
                "Warning | تنبيه",
                "[العربية] يرجى فتح مستند نصي (LibreOffice Writer) أولاً.\n"
                "[English] Please open a text document (LibreOffice Writer) first."
            )
            return

        undo_manager = None
        try:
            if hasattr(doc, "getUndoManager"):
                undo_manager = doc.getUndoManager()
                if undo_manager:
                    undo_manager.enterUndoContext("Format & Align | ضبط التنسيق والمحاذاة")
        except Exception:
            undo_manager = None

        locked = False
        action_locked = False
        try:
            if hasattr(doc, "lockControllers"):
                doc.lockControllers()
                locked = True
            if hasattr(doc, "addActionLock"):
                doc.addActionLock()
                action_locked = True

            is_rtl = self.is_arabic_document(doc)
            self.adjust_footnote_separator(doc, is_rtl)
            self.adjust_footnote_paragraphs(doc, is_rtl)

            fn_settings = doc.FootnoteSettings
            fn_settings.Prefix = "("
            fn_settings.Suffix = ") "
            fn_settings.NumberingType = 4

            fixed = self.fix_existing_footnote_brackets(doc)

            # Unlock document and seal the undo context BEFORE the modal message box
            if action_locked and hasattr(doc, "removeActionLock"):
                try:
                    doc.removeActionLock()
                    action_locked = False
                except Exception:
                    pass
            if locked and hasattr(doc, "unlockControllers"):
                try:
                    doc.unlockControllers()
                    locked = False
                except Exception:
                    pass
            if undo_manager:
                try:
                    undo_manager.leaveUndoContext()
                    undo_manager = None
                except Exception as ex:
                    print(f"[FootnoteConverter] leaveUndoContext error: {ex}")

            self.show_message(
                "Format & Align | ضبط التنسيق والمحاذاة",
                "[ العربية ]\n"
                f"• تم إحاطة ({fixed}) حاشية سفلية بأقواس (1) في المتن عند الحاجة.\n"
                "• تم ضبط ترقيم الحاشية السفلية بصيغة (1).\n"
                "• تم ضبط محاذاة الحواشي وخط الفاصل حسب اتجاه اللغة.\n"
                "• يمكنك التراجع في أي وقت عبر (Ctrl + Z).\n\n"
                "[ English ]\n"
                f"• Enclosed ({fixed}) footnote anchor(s) with parentheses (1) in the text.\n"
                "• Footnote footer numbering set to (1).\n"
                "• Footnotes and separator line are aligned according to the document language.\n"
                "• You can undo at any time using (Ctrl + Z)."
            )
        except Exception as e:
            self.show_message("Error | خطأ", f"Error | خطأ:\n{str(e)}")
        finally:
            if action_locked and hasattr(doc, "removeActionLock"):
                try:
                    doc.removeActionLock()
                except Exception:
                    pass
            if locked and hasattr(doc, "unlockControllers"):
                try:
                    doc.unlockControllers()
                except Exception:
                    pass
            if undo_manager:
                try:
                    undo_manager.leaveUndoContext()
                except Exception:
                    pass

    def unify_footnotes_font(self):
        doc = self.get_current_doc()
        if not doc or not hasattr(doc, "supportsService") or not doc.supportsService("com.sun.star.text.TextDocument"):
            self.show_message(
                "Warning | تنبيه",
                "[العربية] يرجى فتح مستند نصي (LibreOffice Writer) أولاً.\n"
                "[English] Please open a text document (LibreOffice Writer) first."
            )
            return

        undo_manager = None
        try:
            if hasattr(doc, "getUndoManager"):
                undo_manager = doc.getUndoManager()
                if undo_manager:
                    undo_manager.enterUndoContext("Unify Footnotes Font & Size | توحيد خط الحواشي")
        except Exception:
            undo_manager = None

        locked = False
        action_locked = False
        try:
            if hasattr(doc, "lockControllers"):
                doc.lockControllers()
                locked = True
            if hasattr(doc, "addActionLock"):
                doc.addActionLock()
                action_locked = True

            is_rtl = self.is_arabic_document(doc)
            wm = RL_TB if is_rtl else LR_TB
            pa = PARA_RIGHT if is_rtl else PARA_LEFT

            # Detect body text font & size (effective values from the style chain)
            body_font_complex = None
            body_font_western = None
            body_size_complex = 0.0
            body_size_western = 0.0

            style_families = None
            try:
                style_families = doc.getStyleFamilies()
                if style_families and style_families.hasByName("ParagraphStyles"):
                    para_styles = style_families.getByName("ParagraphStyles")
                    for base_style in ["Standard", "Default Paragraph Style", "Body Text"]:
                        if para_styles.hasByName(base_style):
                            bs = para_styles.getByName(base_style)
                            if hasattr(bs, "CharFontNameComplex") and bs.CharFontNameComplex and not body_font_complex:
                                body_font_complex = bs.CharFontNameComplex
                            if hasattr(bs, "CharHeightComplex") and float(bs.CharHeightComplex) > 0 and body_size_complex <= 0:
                                body_size_complex = float(bs.CharHeightComplex)
                            if hasattr(bs, "CharFontName") and bs.CharFontName and not body_font_western:
                                body_font_western = bs.CharFontName
                            if hasattr(bs, "CharHeight") and float(bs.CharHeight) > 0 and body_size_western <= 0:
                                body_size_western = float(bs.CharHeight)
                            if body_size_complex > 0 and body_size_western > 0:
                                break
            except Exception as e:
                print(f"[FootnoteConverter] detect body style error: {e}")

            # Fallback: read the effective font & size of the real body text itself,
            # so that the footnote can never end up larger than the body text
            if body_size_complex <= 0 or body_size_western <= 0:
                try:
                    probe = doc.getText().createTextCursor()
                    probe.gotoStart(False)
                    for _step in range(3):
                        probed = False
                        try:
                            if body_size_complex <= 0 and float(probe.CharHeightComplex) > 0:
                                body_size_complex = float(probe.CharHeightComplex)
                                body_font_complex = body_font_complex or probe.CharFontNameComplex
                                probed = True
                            if body_size_western <= 0 and float(probe.CharHeight) > 0:
                                body_size_western = float(probe.CharHeight)
                                body_font_western = body_font_western or probe.CharFontName
                                probed = True
                        except Exception:
                            pass
                        if probed or not probe.gotoNextParagraph(False):
                            break
                except Exception as e:
                    print(f"[FootnoteConverter] probe body text error: {e}")

            # Academic defaults, only used when nothing at all could be detected
            if not body_font_complex:
                body_font_complex = "Traditional Arabic"
            if not body_font_western:
                body_font_western = "Times New Roman"
            if body_size_complex <= 0:
                body_size_complex = 18.0
            if body_size_western <= 0:
                body_size_western = 12.0

            # Footnote font size = body size - 2 pt (never below the 8 pt floor)
            target_font_complex = body_font_complex
            target_size_complex = max(8.0, body_size_complex - 2.0)
            target_font_western = body_font_western
            target_size_western = max(8.0, body_size_western - 2.0)

            # 1. Update Footnote paragraph style in document
            try:
                if style_families and style_families.hasByName("ParagraphStyles"):
                    para_styles = style_families.getByName("ParagraphStyles")
                    for style_name in ["Footnote", "footnote"]:
                        if para_styles.hasByName(style_name):
                            ps = para_styles.getByName(style_name)
                            try:
                                ps.CharFontNameComplex = target_font_complex
                                ps.CharHeightComplex = target_size_complex
                                ps.CharFontName = target_font_western
                                ps.CharHeight = target_size_western
                                ps.WritingMode = wm
                                ps.ParaAdjust = pa
                            except Exception:
                                pass
            except Exception as e:
                print(f"[FootnoteConverter] update Footnote paragraph style error: {e}")

            # 2. Iterate through all individual footnotes and unify direct formatting
            fns = doc.getFootnotes()
            count = fns.getCount() if fns else 0
            for i in range(count):
                try:
                    fn = fns.getByIndex(i)
                    fn_cur = fn.getText().createTextCursor()
                    fn_cur.gotoStart(False)
                    fn_cur.gotoEnd(True)
                    fn_cur.CharFontNameComplex = target_font_complex
                    fn_cur.CharHeightComplex = target_size_complex
                    fn_cur.CharFontName = target_font_western
                    fn_cur.CharHeight = target_size_western
                    fn_cur.WritingMode = wm
                    fn_cur.ParaAdjust = pa
                except Exception:
                    pass

            # 3. Unlock document and seal the undo context BEFORE the modal message box
            if action_locked and hasattr(doc, "removeActionLock"):
                try:
                    doc.removeActionLock()
                    action_locked = False
                except Exception:
                    pass
            if locked and hasattr(doc, "unlockControllers"):
                try:
                    doc.unlockControllers()
                    locked = False
                except Exception:
                    pass
            if undo_manager:
                try:
                    undo_manager.leaveUndoContext()
                    undo_manager = None
                except Exception:
                    pass

            msg = (
                f"تم بنجاح توحيد خط وحجم الحواشي السفلية! | Success!\n"
                f"-----------------------------------------\n"
                f"[ العربية ]\n"
                f"• نوع الخط الموحد: {target_font_complex}\n"
                f"• حجم خط الحاشية: {target_size_complex:g} نقطة (أصغر بنقطتين من المتن: {body_size_complex:g} نقطة)\n"
                f"• عدد الحواشي الموحدة: ({count}) حاشية.\n"
                f"• تم ضبط اتجاه النص ومحاذاة الفاصل لليمين.\n"
                f"• يمكنك التراجع في أي وقت عبر (Ctrl + Z).\n\n"
                f"[ English ]\n"
                f"• Unified Font: {target_font_complex} / {target_font_western}\n"
                f"• Footnote Size: {target_size_complex:g} pt (Body size: {body_size_complex:g} pt - 2 pt)\n"
                f"• Total Unified Footnotes: ({count})\n"
                f"• Undo supported via (Ctrl + Z)."
            )
            self.show_message("Unify Font & Size | توحيد خط الحواشي", msg)

        except Exception as e:
            self.show_message("Error | خطأ", f"Error | خطأ:\n{str(e)}")
        finally:
            if action_locked and hasattr(doc, "removeActionLock"):
                try:
                    doc.removeActionLock()
                except Exception:
                    pass
            if locked and hasattr(doc, "unlockControllers"):
                try:
                    doc.unlockControllers()
                except Exception:
                    pass
            if undo_manager:
                try:
                    undo_manager.leaveUndoContext()
                except Exception:
                    pass

    def show_about(self):
        msg = (
            "Footnote Converter Pro | محول الحواشي المطور\n"
            "Version | الإصدار: 2.3.1\n"
            "====================================\n\n"
            "[ العربية - المميزات ]\n"
            "• تحويل علامات الحواشي الحصرية < > إلى حواشي معتمدة بدقة.\n"
            "• توافق تام 100% مع Microsoft Word (.docx) عند الحفظ والتصدير للمشرفين.\n"
            "• توحيد خط الحواشي وحجمها تلقائياً (أصغر بنقطتين من المتن: Font Size - 2 pt).\n"
            "• حماية كاملة لأقواس الاقتباس « » وأقواس الحديث (( )) من التحويل.\n"
            "• تجاهل وتطهير الحواشي الفارغة <> ومنع إنشاء حواشي وهمية.\n"
            "• إحاطة النص المحدد بـ < > بضغطة زر واحدة (F2) أو من القائمة.\n"
            "• ترقيم متن النص بين قوسين (1) مع إزالة الفراغ قبل القوس.\n"
            "• ترقيم الحاشية السفلية بالأسفل بصيغة (1).\n"
            "• محاذاة نصوص الحواشي لليمين (RTL) تحت الفاصل مباشرة.\n"
            "• تصحيح اتجاه خط فاصل الحاشية تلقائياً لجهة اليمين للعربية واليسار للإنجليزية.\n"
            "• خيار ترقيم الحواشي لكل صفحة مستقلة أو لكامل المستند.\n"
            "• ميزة التحويل العكسي لاستخراج الحواشي إلى نصوص < >.\n"
            "• تراجع بخطوة واحدة (Ctrl + Z) لكل العمليات: التحويل والتنسيق وتوحيد الخط.\n\n"
            "------------------------------------\n\n"
            "[ English - Features ]\n"
            "• Safely converts footnote markers < > to official footnotes.\n"
            "• 100% compatible with Microsoft Word (.docx) for academic submission.\n"
            "• Unifies footnote font & size automatically (Body size - 2 pt).\n"
            "• Preserves quotes « » and citations (( )) from accidental conversion.\n"
            "• Automatically cleans empty markers <> without creating dummy footnotes.\n"
            "• Wrap selected text in < > via shortcut (F2) or from menu.\n"
            "• Encloses footnote numbers in text in parentheses (1) without gap.\n"
            "• Formats bottom footnote numbering with parentheses: (1).\n"
            "• Automatically aligns separator line and text (RTL for Arabic, LTR for English).\n"
            "• Option to restart footnote numbering per page or continuously.\n"
            "• Reverse conversion: Extract footnotes back to < > text markers.\n"
            "• Full single-step undo (Ctrl + Z) for every operation: convert, format & unify.\n"
            "===================================="
        )
        self.show_message("About | حول محول الحواشي - Footnote Converter Pro v2.3.1", msg)


class FootnoteConverter(unohelper.Base, XJobExecutor, XServiceInfo):
    def __init__(self, ctx):
        self.ctx = ctx
        self.service = FootnoteService(ctx)

    def trigger(self, args):
        cmd = str(args).lower()
        if "about" in cmd:
            self.service.show_about()
        elif "reverse" in cmd:
            self.service.reverse_footnotes()
        elif "page" in cmd:
            self.service.set_numbering_scope(PER_PAGE)
        elif "continuous" in cmd:
            self.service.set_numbering_scope(PER_DOCUMENT)
        elif "wrap" in cmd:
            self.service.wrap_selection_with_markers()
        elif "format" in cmd:
            self.service.format_numbering()
        elif "unify" in cmd or "font" in cmd:
            self.service.unify_footnotes_font()
        else:
            self.service.convert_footnotes()

    def getImplementationName(self):
        return JOB_EXECUTOR_NAME

    def supportsService(self, service_name):
        return service_name in self.getSupportedServiceNames()

    def getSupportedServiceNames(self):
        return (JOB_EXECUTOR_NAME, "com.sun.star.task.JobExecutor")


class FootnoteDispatch(unohelper.Base, XDispatch):
    def __init__(self, ctx, frame):
        self.ctx = ctx
        self.frame = frame
        self.service = FootnoteService(ctx)

    def dispatch(self, url, args):
        cmd = url.Path.lower() if hasattr(url, "Path") else str(url).lower()
        full_url = url.Complete.lower() if hasattr(url, "Complete") else str(url).lower()

        if "about" in cmd or "about" in full_url:
            self.service.show_about()
        elif "numbering_per_page" in cmd or "numbering_per_page" in full_url:
            self.service.set_numbering_scope(PER_PAGE)
        elif "numbering_continuous" in cmd or "numbering_continuous" in full_url:
            self.service.set_numbering_scope(PER_DOCUMENT)
        elif "reverse" in cmd or "reverse" in full_url:
            self.service.reverse_footnotes()
        elif "wrap_selection" in cmd or "wrap_selection" in full_url:
            self.service.wrap_selection_with_markers()
        elif "format_numbering" in cmd or "format_numbering" in full_url:
            self.service.format_numbering()
        elif "unify_font" in cmd or "unify_font" in full_url or "font" in cmd:
            self.service.unify_footnotes_font()
        else:
            self.service.convert_footnotes()

    def addStatusListener(self, listener, url):
        if FeatureStateEvent is not None:
            try:
                state = FeatureStateEvent()
                state.FeatureURL = url
                state.Source = self
                state.IsEnabled = True
                state.Requery = False
                state.State = None
                listener.statusChanged(state)
            except Exception:
                pass

    def removeStatusListener(self, listener, url):
        pass


class FootnoteProtocolHandler(unohelper.Base, XDispatchProvider, XInitialization, XServiceInfo):
    def __init__(self, ctx):
        self.ctx = ctx
        self.frame = None

    def initialize(self, args):
        if args:
            self.frame = args[0]

    def queryDispatch(self, url, target_frame_name, search_flags):
        full_url = url.Complete if hasattr(url, "Complete") else str(url)
        if "vnd.footnote.converter" in full_url or "com.example.footnoteconverter" in full_url:
            return FootnoteDispatch(self.ctx, self.frame)
        return None

    def queryDispatches(self, requests):
        return tuple(
            self.queryDispatch(req.FeatureURL, req.FrameName, req.SearchFlags)
            for req in requests
        )

    def getImplementationName(self):
        return PROTOCOL_HANDLER_NAME

    def supportsService(self, service_name):
        return service_name in self.getSupportedServiceNames()

    def getSupportedServiceNames(self):
        return ("com.sun.star.frame.ProtocolHandler", PROTOCOL_HANDLER_NAME)


# UNO Registration Helper
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    FootnoteConverter,
    JOB_EXECUTOR_NAME,
    (JOB_EXECUTOR_NAME, "com.sun.star.task.JobExecutor"),
)

g_ImplementationHelper.addImplementation(
    FootnoteProtocolHandler,
    PROTOCOL_HANDLER_NAME,
    ("com.sun.star.frame.ProtocolHandler", PROTOCOL_HANDLER_NAME),
)