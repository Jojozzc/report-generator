class RawDataMode2:
    complete_date: str = None

    # 委托单编号
    sample_no: str = None

    # 单线号
    line_no: str = None

    # 焊口编号
    kind_no: str = None

    # 焊工编号
    emp_id: str = None

    # 返修张/处数
    ret_cnt: int = None

    # 材质
    material: str = None

    # 规格
    specification: str = None


    # 检测结果
    check_result: str = None



    def __init__(self,
                 complete_date: str,
                 sample_no: str,
                 line_no: str,
                 kind_no: str,
                 emp_id: str,
                 ret_cnt: int,
                 material: str,
                 specification: str,
                 check_result: str,
                 ):
        self.complete_date = complete_date
        self.sample_no = sample_no
        self.line_no = line_no
        self.kind_no = kind_no
        self.emp_id = emp_id
        self.ret_cnt = ret_cnt
        self.material = material
        self.specification = specification
        self.check_result = check_result
