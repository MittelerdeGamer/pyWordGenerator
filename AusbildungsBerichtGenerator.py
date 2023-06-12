from docx import Document
import datetime


class AusbildungsBericht:
    # $ABJ      Ausbildungsjahr
    __abl: str

    # $ABWB     Ausbildungswoche Beginning
    __abwb: str

    # $ABWE     Ausbildungswoche Ending
    __abwe: str

    # $BT

    # $BTS

    # $U

    # $US

    # $TDB

    # $TDBS

    __doc: Document()
    __ab_begin: datetime

    def __init__(self):
        self.__doc = Document("BerichtVorlage.docx")
        self.__ab_begin = datetime.datetime(2022, 8, 1)

    def __calc_abj(self, abwb: datetime) -> int:
        diff = abwb.year - self.__ab_begin.year
        if abwb.isocalendar()[1] > self.__ab_begin.isocalendar()[1]:
            diff -= 1
        return diff

    def set_head_table(self, kw: str):
        # kw sould be format like jjjj-Www Example: "2023-W26"
        # abwb means Ausbildungswoche Beginning
        # abwe means Ausbildungswoche Ending
        # abj means Ausbildungsjahr
        abwb = datetime.datetime.strptime(kw + '-1', "%Y-W%W-%w")
        abwe = datetime.datetime.strptime(kw + '-5', "%Y-W%W-%w")
        abj = self.__calc_abj(abwb)

    def save_to(self, filename: str):
        for table in self.__doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text
                    if text.startswith("$"):
                        print(text)
        self.__doc.save(filename)
        self.__doc = Document("BerichtVorlage.docx")

    def print_paragraphs(self):
        for p in self.__doc.paragraphs:
            print("-" * 32 + "\nNew Paragraph\n" + "-" * 32)
            print(p.text)

    def print_tables(self):
        for table in self.__doc.tables:
            print("-" * 32 + "\nNew Table\n" + "-" * 32)
            for row in table.rows:
                print("|".join([cell.text for cell in row.cells]))

    def get_paragraphs(self):
        return self.__doc.paragraphs

    def get_tables(self):
        return self.__doc.tables


def main():
    abb = AusbildungsBericht()
    abb.set_head_table("2023-W13")
    abb.print_tables()


if __name__ == '__main__':
    main()
