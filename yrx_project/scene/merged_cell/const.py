from enum import Enum


class LanguageEnum(str, Enum):
    PYTHON = "python"


if __name__ == '__main__':
    print(LanguageEnum.PYTHON == "python")