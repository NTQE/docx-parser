from dataclasses import dataclass, field
import os
import re


@dataclass
class LockoutRow:
    """
    Custom object to reflect the information gathered from each row in the lockout table of the documents
    """
    num: str = ""
    point: str = ""
    tag_no: str = ""
    energy_src: str = ""
    isolating_means: str = ""
    context: bool = False
    context_text: str = ""


@dataclass
class EsipData:
    """
    Custom object to reflect the information gathered from each ESIP document
    """
    abs_path: str
    equip_desc: str = ""
    _doc_no: str = ""
    equip_id: str = ""
    prep_by: str = ""
    prep_by_title: str = ""
    app_by: str = ""
    app_by_title: str = ""
    date_eff: str = ""
    dept: str = ""
    _special_precautions: str = ""
    joined: bool = False
    lockout: list[LockoutRow] = field(default_factory=list)
    extra: list[str] = field(default_factory=list)

    def __post_init__(self):
        self.file_name = os.path.split(self.abs_path)[1]

    @property
    def title(self):
        return f"{self.equip_desc} / {self.equip_id}"

    @property
    def special_precautions(self):
        find = re.search(r"(Special Precautions:)(.*)", self._special_precautions, flags=re.DOTALL)
        if find:
            return find.group(2)
        else:
            return self._special_precautions

    @special_precautions.setter
    def special_precautions(self, value):
        self._special_precautions = value

    @property
    def doc_no(self):
        return f"{self._doc_no}"

    @doc_no.setter
    def doc_no(self, value):
        self._doc_no = f"8LWI.{value}"

    def __repr__(self):
        return f"EsipData(dept={self.dept.ljust(28)[:27]} -- desc={self.equip_desc.ljust(20)[:19]} -- doc={self.doc_no.ljust(15)} -- id={self.equip_id.ljust(15)} -- prep={self.prep_by}/{self.prep_by_title} -- app={self.app_by}/{self.app_by_title} -- date={self.date_eff.ljust(17)},\n\tfile={self.file_name}, \n\tprecautions={self.special_precautions}, \n\tlockout={self.lockout}"
