"""Excel 4.0 (XLM) macro functions, which the formula gate refuses.

Defined names are evaluated with macro rights: a name such as GET.WORKBOOK(1) or FILES("C:\\*")
reads the file system, other workbooks and the application, and RUN, EXEC and CALL start
other code. The functions that act on files, the system, other programs or the workbook's
own structure are blocked as exact names and as families that share a prefix.
"""

_FAMILIES = (
    "GET.",  # GET.WORKBOOK, GET.WINDOW, GET.FORMULA, GET.NAME, GET.DEF, GET.OBJECT, ...
    "APP.",  # APP.TITLE, APP.ACTIVATE, APP.MOVE, ...
    "FILE.",  # FILE.CLOSE, FILE.DELETE, FILE.EXISTS, ...
    "WORKBOOK.",  # WORKBOOK.ADD, WORKBOOK.COPY, WORKBOOK.DELETE, ...
    "ON.",  # ON.TIME, ON.KEY, ON.DATA, ON.WINDOW, ...
    "SEND.",  # SEND.KEYS, SEND.MAIL
    "SET.",  # SET.NAME, SET.VALUE, ...
    "DEFINE.",  # DEFINE.NAME
    "DELETE.",  # DELETE.NAME, DELETE.FORMAT, ...
    "OPEN.",  # OPEN.LINKS, OPEN.TEXT, OPEN.DIALOG
    "SAVE.",  # SAVE.AS, SAVE.COPY.AS, SAVE.WORKSPACE
    "UPDATE.",  # UPDATE.LINK
    "CHANGE.",  # CHANGE.LINK
    "FORMULA.",  # FORMULA.FILL, FORMULA.ARRAY, ...
)

_FUNCTIONS = frozenset(
    {
        # Files and other workbooks
        "FILES",
        "DIRECTORY",
        "DOCUMENTS",
        "WINDOWS",
        "NAMES",
        "LINKS",
        "FOPEN",
        "FCLOSE",
        "FREAD",
        "FREADLN",
        "FWRITE",
        "FWRITELN",
        "FPOS",
        "FSIZE",
        "OPEN",
        "SAVE",
        "NEW",
        "CLOSE",
        # Other programs and code
        "EXEC",
        "EXECUTE",
        "UNREGISTER",
        "INITIATE",
        "POKE",
        "REQUEST",
        "TERMINATE",
        "RUN",
        "HALT",
        "RESTART",
        "EVALUATE",
        # The application and the user
        "ALERT",
        "ECHO",
        "QUIT",
        "ACTIVATE",
        "FORMULA",
    }
)


def is_macro_function(name: str) -> bool:
    """Whether a normalised, upper-case function name is an Excel 4.0 macro function."""
    return name in _FUNCTIONS or name.startswith(_FAMILIES)
