from itertools import chain
from pathlib import Path
import xlwt
import argparse
import sys
import re
sys.path.append(str(Path(__file__).parent))
from excel_helper import OpenpyxlUtil, write_row

class ExcelWalker(object):

    def __init__(self, root: str) -> None:
        self.root = root
        self.excels: dict[str, Path] =  self._scan(root)

    def _scan(self, root_path: str) -> dict[str: Path]:
        excels = {}
        root = Path(root_path)
        for path in chain(*[root.glob("**/[!~]*" + ext) for ext in (".xls", ".xlsx")]):
           
           excels[path.stem] = path
        return excels
    
    def __repr__(self) -> str:
        return "ExcelWalker >>>\n" + "\n".join(map(str, self.excels.values())) + "\n<<<"
    

def run(root_path: Path):
    walker = ExcelWalker(root_path.absolute())
    print(walker)
    header_items = {}
    for name, path in walker.excels.items():
        excel = OpenpyxlUtil(path.absolute(), use_xlrd=path.suffix==".xls")
        for cell in excel.get_sheet_by_index(0).header:
            header_items[format_header(cell)] = 0
    date_pattern = re.compile("\d+")
    data = [["日期"] + list(header_items.keys())]
    for name, path in walker.excels.items():
        excel = OpenpyxlUtil(path.absolute(), use_xlrd=path.suffix==".xls")
        sheet = excel.get_sheet_by_index(0)
        header = [format_header(item) for item in sheet.header]
        line = {"date": "-".join(date_pattern.findall(name))}
        for row in sheet.iter_rows(2):
            line.update(dict.fromkeys(header_items.keys(), ""))
            line.update(dict(zip(header, row)))
            data.append(list(line.values()))
    summary_book = xlwt.Workbook()
    summary_sheet = summary_book.add_sheet("summary")
    for index, line in enumerate(data):
        write_row(summary_sheet, index, line)
    save_path = root_path.joinpath("~summary.xls").absolute()
    if save_path.exists():
        print("删除上一次存在总结文件 %s" % save_path)
        save_path.unlink()
    summary_book.save(save_path)
    print("已完成, 结果保存到 %s" % save_path)


def format_header(header: str) -> str:
    return " ".join([word.capitalize() for word in header.split(" ")])



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="汇总指定目录下所有月薪文件")
    parser.add_argument("root")
    args = parser.parse_args()
    run(Path(args.root))
