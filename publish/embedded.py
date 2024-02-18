"""
pack this project for embeded python
items included:
- source file
- quick excutable bat
- package requirement
- clean up bat
"""


from pathlib import Path
import shutil
import subprocess
from zipfile import ZipFile


class EmbededPacker(object):

    def __init__(self, project_root: str) -> None:
        self.project_root = Path(project_root)
        self.working_dir = self.project_root.joinpath(".temp", "embedded_packing")
        self.src_dir = self.project_root.joinpath("src")
        self.requirement_txt = self.working_dir.joinpath("requirement.txt")
        self.content_archive = "content.zip"
        self.archive_path = self.working_dir.joinpath(self.content_archive)
        self.setup_bat = self.working_dir.joinpath("setup.bat")
        self.release_package = self.project_root.joinpath("release", "gemt")
    
    def prepare_working_dir(self):
        working_dir =  self.working_dir
        if working_dir.exists():
            if working_dir.is_dir():
                self.remove_working_dir()
            else:
                raise Exception(f"{working_dir} 已存在文件")
        working_dir.mkdir()

    def remove_working_dir(self):
        print(f"移除packing文件夹： {self.working_dir}")
        shutil.rmtree(self.working_dir)

    def archive(self):
        with ZipFile(self.archive_path, "w") as zipf:
            for path in self.project_root.glob("*.bat"):
                zipf.write(path.absolute(), path.relative_to(self.project_root))
            for path in self.src_dir.glob("**/*.py"):
                zipf.write(path.absolute(), path.relative_to(self.project_root))

    def gen_requirement(self):
        subprocess.run(f"poetry export --without-hashes --without dev -f requirements.txt -o {self.requirement_txt}", shell=True)


    def build_setup_bat(self):
        with self.setup_bat.open("w") as f:
            f.write("\n".join([
                "RD /S /Q src",
                f"tar -xzvf {self.content_archive}",
                "del /q content.zip setup.bat",
            ]))

    def pack(self, output: Path=None):
        self.prepare_working_dir()
        self.build_setup_bat()
        self.gen_requirement()
        self.archive()
        if output:
            shutil.make_archive(output.with_suffix(""), output.suffix.lstrip("."), self.working_dir)
        else:
            shutil.make_archive(self.release_package, "zip", self.working_dir)



if __name__ == "__main__":
    packer = EmbededPacker(r"D:\Gemt")
    packer.pack(Path(r"D:\Gemt\.temp\setup_test\gemt.zip"))
