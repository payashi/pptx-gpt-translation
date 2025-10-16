from abc import ABC, abstractmethod
from typing import List, Optional


class BaseTranslator(ABC):
    def __init__(
        self,
        model: str,
        source: Optional[str],
        target: str,
        temperature: Optional[float] = None,
    ):
        self.model = model
        self.source = source
        self.target = target
        self.temperature = temperature

    @abstractmethod
    def translate(self, texts: List[str], context: str = "") -> List[str]:
        raise NotImplementedError
