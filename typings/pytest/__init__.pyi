from typing import Any, Callable, TypeVar

_T = TypeVar("_T")


def fixture(*args: Any, **kwargs: Any) -> Callable[[Callable[..., _T]], Callable[..., _T]]:
    ...


def approx(*args: Any, **kwargs: Any) -> Any:
    ...


def __getattr__(name: str) -> Any: ...
