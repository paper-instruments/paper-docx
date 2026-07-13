# pyright: reportPrivateUsage=false

"""In-place rollback for compound paper-docx mutations."""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from typing import (
    TYPE_CHECKING,
    Any,
    Dict,
    Generator,
    List,
    Optional,
    Tuple,
    cast,
)

from lxml import etree

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document
    from docx.opc.package import OpcPackage
    from docx.opc.part import Part
    from docx.opc.rel import Relationships, _Relationship
    from docx.package import ImageParts


_RAW_TAG = etree._Element.tag
_RAW_TEXT = etree._Element.text
_RAW_TAIL = etree._Element.tail


@dataclass(frozen=True)
class _ElementState:
    element: "_Element"
    tag: Any
    attributes: "Tuple[Tuple[str, str], ...]"
    text: Optional[str]
    tail: Optional[str]
    children: "Tuple[_Element, ...]"
    proxy_attributes: "Optional[Dict[str, Any]]"


@dataclass(frozen=True)
class _TreeState:
    elements: "Tuple[_ElementState, ...]"

    @classmethod
    def capture(cls, root: "_Element") -> "_TreeState":
        return cls(
            tuple(
                _ElementState(
                    element=element,
                    tag=_RAW_TAG.__get__(element),
                    attributes=tuple(element.attrib.items()),
                    text=_RAW_TEXT.__get__(element),
                    tail=_RAW_TAIL.__get__(element),
                    children=tuple(element),
                    proxy_attributes=(
                        dict(element.__dict__) if hasattr(element, "__dict__") else None
                    ),
                )
                for element in root.iter()
            )
        )

    def restore(self) -> None:
        # Preserve every unchanged edge. Detaching and re-appending an unchanged
        # node can make lxml reconcile namespace prefixes and invalidate QName
        # values held in ordinary string attributes.
        for state in self.elements:
            self._restore_children(state)

        for state in self.elements:
            element = state.element
            if _RAW_TAG.__get__(element) != state.tag:
                _RAW_TAG.__set__(element, state.tag)
            if tuple(element.attrib.items()) != state.attributes:
                element.attrib.clear()
                for name, value in state.attributes:
                    element.set(name, value)
            if _RAW_TEXT.__get__(element) != state.text:
                _RAW_TEXT.__set__(element, state.text)
            if _RAW_TAIL.__get__(element) != state.tail:
                _RAW_TAIL.__set__(element, state.tail)
            if state.proxy_attributes is not None:
                element.__dict__.clear()
                element.__dict__.update(state.proxy_attributes)

    @staticmethod
    def _restore_children(state: _ElementState) -> None:
        element = state.element
        desired = state.children
        if _same_nodes(tuple(element), desired):
            return
        desired_ids = {id(child) for child in desired}
        for child in tuple(element):
            if id(child) not in desired_ids:
                element.remove(child)
        for index, child in enumerate(desired):
            if child.getparent() is element and element.index(child) == index:
                continue
            element.insert(index, child)
        for child in tuple(element):
            if id(child) not in desired_ids:
                element.remove(child)
        if not _same_nodes(tuple(element), desired):
            raise RuntimeError("package rollback could not restore XML topology")


def _same_nodes(left: "Tuple[_Element, ...]", right: "Tuple[_Element, ...]") -> bool:
    return len(left) == len(right) and all(a is b for a, b in zip(left, right))


@dataclass(frozen=True)
class _ObjectState:
    instance: Any
    attributes: "Dict[str, Any]"
    mutable_values: "Tuple[_MutableValueState, ...]"

    @classmethod
    def capture(cls, instance: Any) -> "_ObjectState":
        attributes = dict(instance.__dict__)
        return cls(
            instance,
            attributes,
            tuple(
                state
                for value in attributes.values()
                if (state := _MutableValueState.capture(value)) is not None
            ),
        )

    def restore(self) -> None:
        for state in self.mutable_values:
            state.restore()
        self.instance.__dict__.clear()
        self.instance.__dict__.update(self.attributes)


@dataclass(frozen=True)
class _MutableValueState:
    value: Any
    kind: str
    items: Any

    @classmethod
    def capture(cls, value: Any) -> "Optional[_MutableValueState]":
        if isinstance(value, list):
            return cls(value, "list", tuple(value))
        if isinstance(value, dict):
            return cls(value, "dict", tuple(value.items()))
        if isinstance(value, set):
            return cls(value, "set", frozenset(value))
        return None

    def restore(self) -> None:
        if self.kind == "list":
            self.value[:] = self.items
        elif self.kind == "dict":
            self.value.clear()
            self.value.update(self.items)
        else:
            self.value.clear()
            self.value.update(self.items)


@dataclass(frozen=True)
class _RelationshipsState:
    relationships: "Relationships"
    base_uri: str
    items: "Tuple[Tuple[str, _Relationship], ...]"
    target_parts: "Dict[str, Any]"
    target_items: "Tuple[Tuple[str, Any], ...]"
    relationship_objects: "Tuple[_ObjectState, ...]"

    @classmethod
    def capture(cls, relationships: "Relationships") -> "_RelationshipsState":
        target_parts = relationships._target_parts_by_rId  # noqa: SLF001
        return cls(
            relationships,
            relationships._baseURI,  # noqa: SLF001
            tuple(relationships.items()),
            target_parts,
            tuple(target_parts.items()),
            tuple(_ObjectState.capture(relationship) for relationship in relationships.values()),
        )

    def restore(self) -> None:
        for state in self.relationship_objects:
            state.restore()
        relationships = cast("Dict[str, _Relationship]", self.relationships)
        relationships.clear()
        relationships.update(self.items)
        self.target_parts.clear()
        self.target_parts.update(self.target_items)
        self.relationships._target_parts_by_rId = self.target_parts  # noqa: SLF001
        self.relationships._baseURI = self.base_uri  # noqa: SLF001

    def is_restored(self) -> bool:
        return (
            _same_items(tuple(self.relationships.items()), self.items)
            and _same_items(tuple(self.target_parts.items()), self.target_items)
            and self.relationships._baseURI == self.base_uri  # noqa: SLF001
            and all(
                state.instance.__dict__ == state.attributes for state in self.relationship_objects
            )
        )


@dataclass(frozen=True)
class _ListState:
    values: "List[Any]"
    items: "Tuple[Any, ...]"

    def restore(self) -> None:
        self.values[:] = self.items

    def is_restored(self) -> bool:
        return _same_objects(tuple(self.values), self.items)


@dataclass(frozen=True)
class _PartBlobState:
    part: "Part"
    blob: bytes

    def is_restored(self) -> bool:
        return self.part.blob == self.blob


@dataclass(frozen=True)
class _PackageState:
    objects: "Tuple[_ObjectState, ...]"
    relationships: "Tuple[_RelationshipsState, ...]"
    trees: "Tuple[_TreeState, ...]"
    image_parts: "Optional[_ListState]"
    part_blobs: "Tuple[_PartBlobState, ...]"

    @classmethod
    def capture(cls, document: "Document") -> "_PackageState":
        package = document.part.package
        assert package is not None
        owners, parts = _reachable_package_objects(package)
        image_parts = cast("Optional[ImageParts]", package.__dict__.get("image_parts"))
        return cls(
            objects=tuple(_ObjectState.capture(owner) for owner in owners),
            relationships=tuple(
                _RelationshipsState.capture(relationships)
                for owner in owners
                if (relationships := cast("Optional[Relationships]", owner.__dict__.get("rels")))
                is not None
            ),
            trees=tuple(
                _TreeState.capture(root)
                for part in parts
                if (root := cast("Optional[_Element]", getattr(part, "_element", None))) is not None
            ),
            image_parts=(
                None
                if image_parts is None
                else _ListState(
                    image_parts._image_parts,  # noqa: SLF001
                    tuple(image_parts._image_parts),  # noqa: SLF001
                )
            ),
            part_blobs=tuple(_PartBlobState(part, part.blob) for part in parts),
        )

    def restore(self) -> None:
        package = cast("OpcPackage", self.objects[0].instance)
        current_parts = _current_package_parts(
            package, tuple(state.instance for state in self.objects)
        )
        original_part_ids = {id(state.part) for state in self.part_blobs}
        # Restore owner attributes before the mutable collaborators they name.
        # This also removes lazy properties created only by the failed edit.
        for state in self.objects:
            state.restore()
        for state in self.relationships:
            state.restore()
        if self.image_parts is not None:
            self.image_parts.restore()
        for state in self.trees:
            state.restore()
        for part in current_parts:
            if id(part) not in original_part_ids and getattr(part, "_package", None) is package:
                part._package = None  # noqa: SLF001 - detach failed-transaction parts
        failures = [
            str(state.part.partname) for state in self.part_blobs if not state.is_restored()
        ]
        if any(not state.is_restored() for state in self.relationships):
            failures.append("<relationships>")
        if self.image_parts is not None and not self.image_parts.is_restored():
            failures.append("<image-parts>")
        if failures:
            raise RuntimeError(
                "package rollback could not restore serialized state: "
                + ", ".join(sorted(failures))
            )


def _same_objects(left: "Tuple[Any, ...]", right: "Tuple[Any, ...]") -> bool:
    return len(left) == len(right) and all(a is b for a, b in zip(left, right))


def _same_items(
    left: "Tuple[Tuple[Any, Any], ...]",
    right: "Tuple[Tuple[Any, Any], ...]",
) -> bool:
    return len(left) == len(right) and all(
        left_key == right_key and left_value is right_value
        for (left_key, left_value), (right_key, right_value) in zip(left, right)
    )


def _reachable_package_objects(
    package: "OpcPackage",
) -> "Tuple[Tuple[OpcPackage | Part, ...], Tuple[Part, ...]]":
    """Return every package-owned part without materializing lazy state."""
    owners: "List[OpcPackage | Part]" = [package]
    parts: "List[Part]" = []
    seen: "set[int]" = set()

    def add(part: "Part") -> None:
        if id(part) in seen:
            return
        seen.add(id(part))
        parts.append(part)
        owners.append(part)

    for owner in owners:
        relationships = cast("Optional[Relationships]", owner.__dict__.get("rels"))
        if relationships is not None:
            for relationship in relationships.values():
                if not relationship.is_external:
                    add(relationship.target_part)
            for part in relationships._target_parts_by_rId.values():  # noqa: SLF001
                add(cast("Part", part))
        if owner is package:
            image_parts = cast("Optional[ImageParts]", package.__dict__.get("image_parts"))
            if image_parts is not None:
                for part in image_parts._image_parts:  # noqa: SLF001
                    add(part)
    return tuple(owners), tuple(parts)


def _current_package_parts(package: "OpcPackage", seeds: "Tuple[Any, ...]") -> "Tuple[Part, ...]":
    """Best-effort raw-reference scan tolerating a corrupted live graph."""
    owners = list(seeds)
    parts: "List[Part]" = []
    seen = set()

    def add(candidate: Any) -> None:
        if id(candidate) in seen:
            return
        attributes = getattr(candidate, "__dict__", None)
        if not isinstance(attributes, dict) or "_package" not in attributes:
            return
        seen.add(id(candidate))
        parts.append(cast("Part", candidate))
        owners.append(candidate)

    for owner in owners:
        attributes = getattr(owner, "__dict__", {})
        relationships = attributes.get("rels")
        if isinstance(relationships, dict):
            for relationship in tuple(relationships.values()):
                rel_attributes = getattr(relationship, "__dict__", {})
                if not rel_attributes.get("_is_external", False):
                    add(rel_attributes.get("_target"))
            target_index = getattr(relationships, "_target_parts_by_rId", None)
            if isinstance(target_index, dict):
                for candidate in tuple(target_index.values()):
                    add(candidate)
        if owner is package:
            image_parts = attributes.get("image_parts")
            image_list = getattr(image_parts, "_image_parts", None)
            if isinstance(image_list, list):
                for candidate in tuple(image_list):
                    add(candidate)
    return tuple(parts)


@contextmanager
def rollback_on_error(document: "Document", *participants: Any) -> Generator[None, None, None]:
    """Restore the live package and named mutable proxies after an error."""
    state = _PackageState.capture(document)
    participant_states = tuple(_ObjectState.capture(participant) for participant in participants)
    try:
        yield
    except BaseException:
        try:
            state.restore()
        finally:
            for participant_state in participant_states:
                participant_state.restore()
        raise


@contextmanager
def rollback_xml_on_error(root: "_Element") -> Generator[None, None, None]:
    """Restore one live XML tree after an error in a local compound edit."""
    state = _TreeState.capture(root)
    before = etree.tostring(root)
    try:
        yield
    except BaseException:
        state.restore()
        if etree.tostring(root) != before:
            raise RuntimeError("XML rollback could not restore serialized state")
        raise
