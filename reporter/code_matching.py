from abc import ABC, abstractmethod
from reporter import common
import re


class Filter(ABC):
    @abstractmethod
    def filter(self, program: str, workplace: str) -> bool:
        """ Определяет выполняется ли условие"""


class FilterIntersection(Filter):
    def __init__(self, filter_set, separators: [str]):
        super().__init__()
        self.filter_set = filter_set
        self.separators = separators

    def filter(self, program, workplace) -> bool:
        del workplace
        return set(self._separate(program.lower().strip())) & self.filter_set

    def _separate(self, value: str):
        return re.split('|'.join(self.separators), value)


class FilterWorkplaceExact(Filter):
    def __init__(self, exact):
        super().__init__()
        self.exact = exact

    def filter(self, program, workplace) -> bool:
        del program
        return workplace.lower().strip() == self.exact.lower()


class FilterInclude(Filter):
    def __init__(self, include, separators: [str]):
        super().__init__()
        self.include = include
        self.separators = separators

    def filter(self, program, workplace) -> bool:
        del workplace
        return self.include in self._separate(program.lower().strip())

    def _separate(self, value: str):
        return re.split('|'.join(self.separators), value)


class Handler(ABC):
    @abstractmethod
    def filter(self, program, workplace) -> bool:
        """ Определить выполняется ли условие для handler """

    @abstractmethod
    def handle(self, access_type: common.AccessTypes):
        """ Обрабатать случай """


class HandlerWithFilters(Handler):
    def __init__(self):
        self.filters: [Filter] = []

    def add_filter(self, _filter: Filter):
        self.filters.append(_filter)

    def filter(self, program, workplace) -> bool:
        return all([f.filter(program, workplace) for f in self.filters])

    def handle(self, access_type: common.AccessTypes):
        pass


class CodeHandler(HandlerWithFilters):
    def __init__(self, access_types_map):
        super().__init__()
        self.access_types_map = access_types_map

    def handle(self, access_type: common.AccessTypes):
        return self.access_types_map.get(access_type)


class CodeMatcher:
    def __init__(self):
        self.handlers: [Handler] = []

    def add_handler(self, handler: Handler):
        self.handlers.append(handler)

    def get_code(self, program, workspace, access_type):
        for handler in self.handlers:
            if handler.filter(program, workspace):
                return handler.handle(access_type)
        return None
