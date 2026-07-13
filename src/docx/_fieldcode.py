"""Minimal, position-preserving parser for bookmark-bearing Word field codes."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Mapping, Optional, Tuple

from docx.errors import UnsupportedStructureError

_FIELD_COMMANDS = frozenset(
    """
    ADDRESSBLOCK ADVANCE ASK AUTHOR AUTONUM AUTONUMLGL AUTONUMOUT AUTOTEXT
    AUTOTEXTLIST BARCODE BIBLIOGRAPHY BIDIOUTLINE CITATION COMMENTS COMPARE
    CREATEDATE DATABASE DATE DISPLAYBARCODE DOCVARIABLE DOCPROPERTY EDITTIME EMBED EQ
    FILENAME FILESIZE FILLIN FORMCHECKBOX FORMDROPDOWN FORMTEXT GLOSSARY
    GOTOBUTTON GREETINGLINE HYPERLINK IF INCLUDEPICTURE INCLUDETEXT INDEX INFO
    KEYWORDS LASTSAVEDBY LINK LISTNUM MACROBUTTON MERGEBARCODE MERGEFIELD
    MERGEREC MERGESEQ NEXT NEXTIF NOTEREF NUMCHARS NUMPAGES NUMWORDS PAGE
    PAGEREF PRINT PRINTDATE PRIVATE QUOTE RD REF REVNUM SAVEDATE SECTION
    SECTIONPAGES SEQ SET SKIPIF STYLEREF SUBJECT SYMBOL TA TC TEMPLATE TIME
    TITLE TOA TOC USERADDRESS USERINITIALS USERNAME XE
    """.split()  # noqa: SIM905 - compact explicit vocabulary
)


@dataclass(frozen=True)
class FieldOperand:
    value: str
    start: int
    end: int
    quoted: bool

    def render(self, value: str) -> str:
        quoted = self.quoted or any(character.isspace() for character in value)
        if not quoted:
            return value
        escaped = value.replace('"', '\\"')
        return f'"{escaped}"'


def _tokens(instruction: str) -> "Tuple[FieldOperand, ...]":
    tokens: list[FieldOperand] = []
    index = 0
    while index < len(instruction):
        while index < len(instruction) and instruction[index].isspace():
            index += 1
        if index >= len(instruction):
            break
        start = index
        if instruction[index] != '"':
            while index < len(instruction) and not instruction[index].isspace():
                index += 1
            tokens.append(FieldOperand(instruction[start:index], start, index, False))
            continue
        index += 1
        value: list[str] = []
        while index < len(instruction):
            char = instruction[index]
            if char == '"':
                index += 1
                tokens.append(FieldOperand("".join(value), start, index, True))
                break
            if char == "\\" and index + 1 < len(instruction) and instruction[index + 1] == '"':
                value.append('"')
                index += 2
                continue
            value.append(char)
            index += 1
        else:
            raise UnsupportedStructureError(
                "field instruction contains an unterminated quoted operand; nothing was changed"
            )
    return tuple(tokens)


def command_operand(instruction: str, command: str) -> "Optional[FieldOperand]":
    tokens = _tokens(instruction)
    if (
        not tokens
        or tokens[0].quoted
        or tokens[0].value.upper() != command.upper()
        or len(tokens) < 2
    ):
        return None
    operand = tokens[1]
    return None if not operand.quoted and operand.value.startswith("\\") else operand


def bookmark_operands(
    instruction: str, *, implicit_names: "Iterable[str]" = ()
) -> "Tuple[FieldOperand, ...]":
    tokens = _tokens(instruction)
    command = "" if not tokens or tokens[0].quoted else tokens[0].value.upper()
    if instruction.lstrip().startswith("="):
        return _formula_bookmark_operands(instruction, implicit_names)
    if command in ("REF", "PAGEREF", "NOTEREF", "GOTOBUTTON"):
        operand = command_operand(instruction, command)
        return (operand,) if operand is not None else ()
    switch = "\\l" if command == "HYPERLINK" else "\\b" if command == "TOC" else None
    if switch:
        return tuple(
            tokens[index + 1]
            for index, token in enumerate(tokens[:-1])
            if not token.quoted
            and token.value.casefold() == switch
            and (tokens[index + 1].quoted or not tokens[index + 1].value.startswith("\\"))
        )
    implicit = {name.casefold() for name in implicit_names}
    if (
        tokens
        and not tokens[0].quoted
        and command not in _FIELD_COMMANDS
        and tokens[0].value.casefold() in implicit
    ):
        return (tokens[0],)
    return ()


def _formula_bookmark_operands(
    instruction: str, names: "Iterable[str]"
) -> "Tuple[FieldOperand, ...]":
    """Known bookmark identifiers in a Word formula, excluding functions/literals."""
    wanted = {name.casefold() for name in names}
    operands: list[FieldOperand] = []
    index = 0
    quoted = False
    while index < len(instruction):
        character = instruction[index]
        if character == '"':
            quoted = not quoted
            index += 1
            continue
        if quoted or not (character.isalpha() or character == "_"):
            index += 1
            continue
        start = index
        index += 1
        while index < len(instruction) and (
            instruction[index].isalnum() or instruction[index] == "_"
        ):
            index += 1
        value = instruction[start:index]
        following = index
        while following < len(instruction) and instruction[following].isspace():
            following += 1
        if (
            value.casefold() in wanted
            and (following >= len(instruction) or instruction[following] != "(")
        ):
            operands.append(FieldOperand(value, start, index, False))
    return tuple(operands)


def rewrite_bookmark_operands(instruction: str, renames: "Mapping[str, str]") -> str:
    lookup = {old.casefold(): new for old, new in renames.items()}
    replacements = [
        (operand, lookup[operand.value.casefold()])
        for operand in bookmark_operands(instruction, implicit_names=renames)
        if operand.value.casefold() in lookup
    ]
    return _rewrite(instruction, replacements)


def rewrite_command_operand(instruction: str, command: str, renames: "Mapping[str, str]") -> str:
    operand = command_operand(instruction, command)
    lookup = {old.casefold(): new for old, new in renames.items()}
    if operand is None or operand.value.casefold() not in lookup:
        return instruction
    return _rewrite(instruction, [(operand, lookup[operand.value.casefold()])])


def _rewrite(
    instruction: str, replacements: "Iterable[tuple[FieldOperand, str]]"
) -> str:
    for operand, replacement in sorted(replacements, key=lambda item: item[0].start, reverse=True):
        instruction = (
            instruction[: operand.start] + operand.render(replacement) + instruction[operand.end :]
        )
    return instruction
