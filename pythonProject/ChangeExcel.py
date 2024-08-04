import openpyxl
import xmind
from xmind.core.markerref import MarkerId
from xmind.core.workbook import WorkbookDocument
# 优先级映射关系
priority_mapping = {
    '高': 'priority-1',
    '中': 'priority-2',
    '低': 'priority-3'
}

def load_excel(file_path):
    """加载Excel并转换为字典"""
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    headers = [cell.value for cell in sheet[1]]
    cases = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        case = dict(zip(headers, row))
        cases.append(case)
    return cases

def create_xmind(cases, output_path):
    """从用例列表创建Xmind文件"""
    workbook = WorkbookDocument()
    root_topic = workbook.getPrimarySheet()
    root_topic.setTitle("测试版本")
    # 创建一个主题
    root_topic = workbook.getPrimarySheet().getRootTopic()
    root_topic.setTitle("我的主题")

    # 保存Xmind文件
    xmind_content = workbook.serialize()
    with open(output_path, 'wb') as f:
        f.write(xmind_content)

if __name__ == '__main__':
    excel_cases = load_excel("D:\Project\\xmind转excel用例\\测试版本1.xlsx")
    create_xmind(excel_cases, r"D:\Project\excel转xmind用例\测试用例.xmind")
