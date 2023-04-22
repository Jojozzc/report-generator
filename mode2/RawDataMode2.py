class RawDataMode2:
    complete_date: str = None

    # 检件编号
    sample_no: str = None

    # 焊口编号
    kind_no: str = None

    # 材质
    material: str = None

    # 规格
    specification: str = None

    # 底片规格
    sample_specification: str = None

    # 底片数量
    sample_cnt: int = None

    # 合格数量
    qualified_sample_cnt: int = None

    # 合格级别
    quality_level: str = None

    def __init__(self,
                 complete_date: str,
                 sample_no: str,
                 kind_no: str,
                 material: str,
                 specification: str,
                 sample_specification: str,
                 sample_cnt: int,
                 qualified_sample_cnt: int,
                 quality_level: str):
        self.complete_date = complete_date
        self.sample_no = sample_no
        self.kind_no = kind_no
        self.material = material
        self.specification = specification
        self.sample_specification = sample_specification
        self.sample_cnt = sample_cnt
        self.qualified_sample_cnt = qualified_sample_cnt
        self.quality_level = quality_level
