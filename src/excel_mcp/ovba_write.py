"""Writing a VBA project (vbaProject.bin), following [MS-OVBA].

The project is rebuilt from source text: every module is stored as compressed
source only, with no p-code, and ``_VBA_PROJECT`` is reset to its
version-independent form, so Office recompiles the macros when it opens the file.
Everything else in the project (references, forms, protection settings) is kept.
"""

import uuid
from dataclasses import dataclass

from excel_mcp import cfb, ovba
from excel_mcp.errors import InvalidArgumentError
from excel_mcp.ovba import ModuleKind, Record
from excel_mcp.ovba_compress import compress

_PROJECT_COOKIE = 0x0013
_PROJECT_END = 0x0010
_MODULE_NON_PROCEDURAL = 0x0022
_VERSION_INDEPENDENT_PROJECT = bytes.fromhex("cc61ffff000000")
_MODULE_LINE_KEYS = ("Module", "Class", "Document", "BaseClass", "Package")
_NEW_CODE_PAGE = 1252
_NEW_WORKSPACE_ENTRY = "0, 0, 0, 0, C"

_CLASS_ATTRIBUTES = (
    'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = False\r\n"
    "Attribute VB_Exposed = False\r\n"
    "Attribute VB_TemplateDerived = False\r\n"
    "Attribute VB_Customizable = False\r\n"
)
_DOCUMENT_ATTRIBUTES = (
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = True\r\n"
    "Attribute VB_Exposed = True\r\n"
    "Attribute VB_TemplateDerived = False\r\n"
    "Attribute VB_Customizable = True\r\n"
)

# Reference records of the two type libraries every Excel project uses: stdole and Office.
_NEW_HEADER: list[Record] = [
    (0x01, (1).to_bytes(4, "little")),
    (0x02, (0x409).to_bytes(4, "little")),
    (0x14, (0x409).to_bytes(4, "little")),
    (0x03, _NEW_CODE_PAGE.to_bytes(2, "little")),
    (0x04, b"VBAProject"),
    (0x05, b""),
    (0x40, b""),
    (0x06, b""),
    (0x3D, b""),
    (0x07, bytes(4)),
    (0x08, bytes(4)),
    (0x09, bytes.fromhex("511f67522000")),
    (0x0C, b""),
    (0x3C, b""),
]


def _registered_reference(name: str, libid: str) -> list[Record]:
    encoded = libid.encode("ascii")
    value = len(encoded).to_bytes(4, "little") + encoded + bytes(4) + bytes(2)
    return [
        (0x16, name.encode("ascii")),
        (0x3E, name.encode("utf-16-le")),
        (0x0D, value),
    ]


_NEW_HEADER += _registered_reference(
    "stdole",
    r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0"
    r"#C:\Windows\System32\stdole2.tlb#OLE Automation",
)
_NEW_HEADER += _registered_reference(
    "Office",
    r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files"
    r"\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library",
)


@dataclass
class Module:
    name: str
    stream: str
    kind: ModuleKind
    source: str
    records: list[Record]


class Project:
    """A VBA project held as text, ready to be changed and serialized again."""

    def __init__(
        self,
        code_page: int,
        header: list[Record],
        text: str,
        modules: list[Module],
        preserved: list[cfb.Entry],
    ) -> None:
        self.code_page = code_page
        self.header = header
        self.text = text
        self.modules = modules
        self.preserved = preserved

    @classmethod
    def from_bytes(cls, content: bytes) -> "Project":
        data = ovba.read_project(content)
        modules = [
            Module(raw.name, raw.stream, raw.kind, raw.source, raw.records) for raw in data.modules
        ]
        replaced = {"VBA/dir", "VBA/_VBA_PROJECT", "PROJECT", "PROJECTwm"}
        replaced |= {f"VBA/{module.stream}" for module in modules}
        preserved = [
            entry
            for entry in cfb.read_entries(content)
            if entry.path not in replaced and not entry.path.startswith("VBA/__SRP_")
        ]
        return cls(data.code_page, data.header, data.text, modules, preserved)

    @classmethod
    def new(cls, documents: list[tuple[str, str]]) -> "Project":
        """A project with one empty document module per (name, VB_Base class id) pair."""
        modules = [
            Module(
                name,
                name,
                "document",
                _attributes(name) + f'Attribute VB_Base = "0{base}"\r\n' + _DOCUMENT_ATTRIBUTES,
                _module_records(name, "document"),
            )
            for name, base in documents
        ]
        text = [f'ID="{{{str(uuid.uuid4()).upper()}}}"']
        text += [f"Document={name}/&H00000000" for name, _ in documents]
        text += [
            'Name="VBAProject"',
            'HelpContextID="0"',
            'VersionCompatible32="393222000"',
            "",
            "[Host Extender Info]",
            "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
            "",
            "[Workspace]",
        ]
        text += [f"{name}={_NEW_WORKSPACE_ENTRY}" for name, _ in documents]
        return cls(_NEW_CODE_PAGE, _NEW_HEADER, _join_lines(text), modules, [])

    def find(self, name: str) -> Module | None:
        return next((m for m in self.modules if m.name.casefold() == name.casefold()), None)

    def set_code(self, name: str, code: str, kind: ModuleKind) -> bool:
        """Replace a module's code, creating the module if needed. Returns True if created."""
        body = code.replace("\r\n", "\n").replace("\r", "\n").strip("\n").replace("\n", "\r\n")
        self._check_encodable(body)
        existing = self.find(name)
        header = (
            _leading_attributes(existing.source)
            if existing
            else _attributes(name) + (_CLASS_ATTRIBUTES if kind == "class" else "")
        )
        source = header + (body + "\r\n" if body else "")
        if existing:
            existing.source = source
            return False
        self.modules.append(Module(name, name, kind, source, _module_records(name, kind)))
        self._add_text_line(f"{'Module' if kind == 'standard' else 'Class'}={name}", name)
        return True

    def remove(self, module: Module) -> None:
        self.modules.remove(module)
        lines = []
        in_workspace = False
        for line in self.text.splitlines():
            key, _, value = line.partition("=")
            in_workspace = in_workspace or line == "[Workspace]"
            is_module_line = (
                key in ("Module", "Class") and value.casefold() == module.name.casefold()
            )
            is_workspace_line = in_workspace and key.casefold() == module.name.casefold()
            if not (is_module_line or is_workspace_line):
                lines.append(line)
        self.text = _join_lines(lines)

    def to_bytes(self) -> bytes:
        encoding = self._encoding
        entries = [
            *self.preserved,
            cfb.Entry("VBA/dir", compress(self._directory())),
            cfb.Entry("VBA/_VBA_PROJECT", _VERSION_INDEPENDENT_PROJECT),
            cfb.Entry("PROJECT", self.text.encode(encoding)),
            cfb.Entry("PROJECTwm", self._module_names()),
        ]
        entries += [
            cfb.Entry(f"VBA/{module.stream}", compress(module.source.encode(encoding)))
            for module in self.modules
        ]
        return cfb.write(entries)

    @property
    def _encoding(self) -> str:
        return ovba.encoding(self.code_page)

    def _check_encodable(self, text: str) -> None:
        try:
            text.encode(self._encoding)
        except UnicodeEncodeError as error:
            raise InvalidArgumentError(
                f"The code contains {error.object[error.start]!r}, which the project's code page "
                f"({self._encoding}) cannot store. Write such characters with ChrW() instead."
            ) from None

    def _directory(self) -> bytes:
        records = [
            *self.header,
            (ovba.PROJECT_MODULES, len(self.modules).to_bytes(2, "little")),
            (_PROJECT_COOKIE, b"\xff\xff"),
        ]
        for module in self.modules:
            records += [
                (record_id, bytes(4) if record_id == ovba.MODULE_OFFSET else value)
                for record_id, value in module.records
            ]
        records.append((_PROJECT_END, b""))
        return b"".join(_encode_record(record_id, value) for record_id, value in records)

    def _module_names(self) -> bytes:
        names = b"".join(
            module.name.encode(self._encoding) + b"\0" + module.name.encode("utf-16-le") + b"\0\0"
            for module in self.modules
        )
        return names + b"\0\0"

    def _add_text_line(self, line: str, workspace_key: str) -> None:
        lines = self.text.splitlines()
        last_module = max(
            (i for i, text in enumerate(lines) if text.partition("=")[0] in _MODULE_LINE_KEYS),
            default=0,
        )
        lines.insert(last_module + 1, line)
        if "[Workspace]" in lines:
            lines.append(f"{workspace_key}={_NEW_WORKSPACE_ENTRY}")
        self.text = _join_lines(lines)


def _attributes(name: str) -> str:
    return f'Attribute VB_Name = "{name}"\r\n'


def _leading_attributes(source: str) -> str:
    """The attribute lines a module's source starts with, which the VBA editor hides."""
    header = ""
    for line in source.split("\r\n"):
        if not line.startswith("Attribute "):
            break
        header += line + "\r\n"
    return header


def _module_records(name: str, kind: ModuleKind) -> list[Record]:
    return [
        (ovba.MODULE_NAME, name.encode("ascii")),
        (0x0047, name.encode("utf-16-le")),
        (0x001A, name.encode("ascii")),
        (0x0032, name.encode("utf-16-le")),
        (0x001C, b""),
        (0x0048, b""),
        (ovba.MODULE_OFFSET, bytes(4)),
        (0x001E, bytes(4)),
        (0x002C, b"\xff\xff"),
        (ovba.MODULE_TYPE_PROCEDURAL if kind == "standard" else _MODULE_NON_PROCEDURAL, b""),
        (ovba.MODULE_END, b""),
    ]


def _encode_record(record_id: int, value: bytes) -> bytes:
    size = 4 if record_id == ovba.PROJECT_VERSION else len(value)
    return record_id.to_bytes(2, "little") + size.to_bytes(4, "little") + value


def _join_lines(lines: list[str]) -> str:
    return "\r\n".join(lines) + "\r\n"
