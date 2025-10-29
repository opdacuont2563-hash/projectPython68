"""Caching helpers for SurgiBot."""
from __future__ import annotations

from functools import lru_cache
from typing import Callable, Iterable, List


def cached_lookup(maxsize: int = 256) -> Callable[[Callable[..., List[str]]], Callable[..., List[str]]]:
    """Decorate a function returning list-like results with LRU caching."""

    def decorator(func: Callable[..., List[str]]) -> Callable[..., List[str]]:
        cached = lru_cache(maxsize=maxsize)(func)

        def wrapper(*args, **kwargs) -> List[str]:
            return list(cached(*args, **kwargs))

        wrapper.cache_clear = cached.cache_clear  # type: ignore[attr-defined]
        return wrapper

    return decorator


def prefix_match(query: str, items: Iterable[str]) -> List[str]:
    q = query.lower().strip()
    if not q:
        return list(items)
    return [item for item in items if item.lower().startswith(q)]


def contains_match(query: str, items: Iterable[str]) -> List[str]:
    q = query.lower().strip()
    if not q:
        return list(items)
    return [item for item in items if q in item.lower()]


__all__ = ["cached_lookup", "prefix_match", "contains_match"]
