from awe_report_generator.core.base import ReportGenerator


class ProcessorUIConfig:
    def __init__(self, name: str, processor: ReportGenerator):
        self.name = name
        self.processor = processor
