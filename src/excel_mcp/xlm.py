"""Excel 4.0 (XLM) macro functions, which the formula gate refuses.

Defined names are evaluated with macro rights: a name such as GET.WORKBOOK(1) or FILES("C:\\*")
reads the file system, other workbooks and the application, and RUN, EXEC and CALL start
other code. The functions of these families are blocked, as exact names and as prefixes:
file system, system and application, external programs and mail, DDE and links, databases,
execution, and reading the structure of the workbook.

The list follows the index of Microsoft's "Excel 4.0 Macro Functions Reference" (527 names),
and the names were checked against Excel on a macro sheet. Functions that only format,
select or edit cells are not blocked, and neither are names that are also worksheet
functions (ROW, COLUMN, INDEX, SORT, FILTER, NOW, CHAR, ...).
"""

_FAMILIES = (
    "GET.",  # GET.WORKBOOK, GET.WINDOW, GET.FORMULA, GET.NAME, GET.DEF, GET.OBJECT, ...
    "APP.",  # APP.TITLE, APP.ACTIVATE, APP.MOVE, ...
    "FILE.",  # FILE.CLOSE, FILE.DELETE, FILE.EXISTS, ...
    "WORKBOOK.",  # WORKBOOK.ADD, WORKBOOK.COPY, WORKBOOK.DELETE, ...
    "ON.",  # ON.TIME, ON.KEY, ON.DATA, ON.WINDOW, ...
    "SEND.",  # SEND.KEYS, SEND.MAIL
    "MAIL.",  # MAIL.LOGON, MAIL.SEND.MAILER, ...
    "SQL.",  # SQL.OPEN, SQL.EXEC.QUERY, ...
    "QUERY.",  # QUERY.GET.DATA, QUERY.REFRESH
    "SOLVER.",  # SOLVER.LOAD, SOLVER.SAVE, ...
    "REPORT.",  # REPORT.GET, REPORT.PRINT, ...
    "VBA.",  # VBA.INSERT.FILE, VBA.MAKE.ADDIN
    "SET.",  # SET.NAME, SET.VALUE, ...
    "DEFINE.",  # DEFINE.NAME
    "DELETE.",  # DELETE.NAME, DELETE.FORMAT, ...
    "OPEN.",  # OPEN.LINKS, OPEN.TEXT, OPEN.DIALOG, OPEN.MAIL
    "SAVE.",  # SAVE.AS, SAVE.COPY.AS, SAVE.WORKSPACE
    "UPDATE.",  # UPDATE.LINK
    "CHANGE.",  # CHANGE.LINK
    "FORMULA.",  # FORMULA.FILL, FORMULA.ARRAY, ...
    "REGISTER.",  # REGISTER.ID
    "SOUND.",  # SOUND.PLAY, SOUND.NOTE
    "ACTIVATE.",  # ACTIVATE.NEXT, ACTIVATE.PREV
)

_FUNCTIONS = frozenset(
    {
        # File system and other workbooks
        "FILES",
        "DIRECTORY",
        "DOCUMENTS",
        "WINDOWS",
        "NAMES",
        "LIST.NAMES",
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
        "CLOSE.ALL",
        "PRINT",
        "LINE.PRINT",
        "INSERT.PICTURE",
        "INSERT.OBJECT",
        "EMBED",
        "SUMMARY.INFO",
        "ADDIN.MANAGER",
        # DDE and links
        "LINKS",
        "INITIATE",
        "POKE",
        "REQUEST",
        "TERMINATE",
        "PASTE.LINK",
        "PASTE.PICTURE.LINK",
        "SUBSCRIBE.TO",
        "CREATE.PUBLISHER",
        "EDITION.OPTIONS",
        "ROUTE.DOCUMENT",
        "ROUTING.SLIP",
        "CLEAR.ROUTING.SLIP",
        # Execution
        "EXEC",
        "EXECUTE",
        "CALL",
        "REGISTER",
        "UNREGISTER",
        "RUN",
        "HALT",
        "RESTART",
        "EVALUATE",
        # The application and the user
        "ALERT",
        "ECHO",
        "QUIT",
        "ACTIVATE",
        "HELP",
        "SHOW.CLIPBOARD",
        "MACRO.OPTIONS",
        # Reading the structure of the workbook
        "SCENARIO.GET",
        "VIEW.GET",
        "SLIDE.GET",
        "OPTIONS.LIST.GET",
        "OPTIONS.LISTS.GET",
        "FORMULA",
    }
)


def is_macro_function(name: str) -> bool:
    """Whether a normalised, upper-case function name is an Excel 4.0 macro function."""
    return name in _FUNCTIONS or name.startswith(_FAMILIES)
