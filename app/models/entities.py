from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Client:
    id: int
    name: str
    location: str = ""


@dataclass(frozen=True, slots=True)
class Product:
    id: int
    name: str
    product_type: str = ""


@dataclass(frozen=True, slots=True)
class Tag:
    id: int
    name: str
    color: str = "#0f766e"

