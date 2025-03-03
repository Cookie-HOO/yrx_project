import typing

import pandas as pd

from yrx_project.const import *


def get_yaml_conf(scene: str, conf_key=None):
    import yaml
    conf_path = os.path.join(CONF_PATH, scene+".yaml")
    with open(conf_path) as f:
        result = yaml.safe_load(f.read())
    if conf_key is None:
        return result
    return result.get(conf_key)


def set_yaml_conf(scene: str, conf_key: str, conf_value):
    import yaml
    conf_path = os.path.join(CONF_PATH, scene+".yaml")
    conf_data = get_yaml_conf(scene)
    conf_data[conf_key] = conf_value
    with open(conf_path, 'w') as yaml_file:
        yaml_file.write(yaml.dump(conf_data, default_flow_style=False, allow_unicode=True))


def get_txt_conf(path, type_=str):
    if type_ == str:
        with open(path, encoding="utf8") as f:
            result = f.read()
        return result
    elif type_ == list:
        with open(path, encoding="utf8") as f:
            result = [i.strip("\n") for i in f.readlines() if i.strip("\n")]
        return result


def set_txt_conf(path, value):
    with open(path, "w", encoding="utf8") as f:
        f.write(value)


def get_csv_conf(path):
    return pd.read_csv(path, encoding="utf-8")


def set_csv_conf(path, value: pd.DataFrame):
    value.to_csv(path, encoding="utf-8", index=False)


class CSVConf:
    def __init__(self, path, init_columns: list = None):
        self.path = path
        self.init_columns = init_columns
        self.df = self.open_or_create()

    def open_or_create(self):
        if os.path.exists(self.path):
            df = pd.read_csv(self.path, encoding="utf-8")
        else:
            if self.init_columns is None:
                raise ValueError("No csv or init_columns")
            df = pd.DataFrame(columns=self.init_columns)
        return df

    def save(self):
        self.set(self.df)

    def get(self):
        return self.df

    def set(self, value: pd.DataFrame):
        self.df = value
        value.to_csv(self.path, encoding="utf-8", index=False)

    def clear(self):
        self.df = self.df[0:0]
        return self

    def append(self, value: typing.Union[pd.DataFrame, typing.List[dict], dict]):
        # 如果 value 是字典，将其转换为 DataFrame
        if isinstance(value, dict):
            value = pd.DataFrame([value])
        # 如果 value 是字典列表，将其转换为 DataFrame
        elif isinstance(value, list) and all(isinstance(item, dict) for item in value):
            value = pd.DataFrame(value)
        # 使用 pd.concat 拼接 DataFrame
        self.df = pd.concat([self.df, value], ignore_index=True)
        return self

    def dedup(self):
        self.df = self.df.drop_duplicates()
        return self


if __name__ == '__main__':
    # print(get_yaml_conf("system", "window"))
    # set_yaml_conf("system", "developer", "fuyao.cookiee")
    # print(get_yaml_conf("system", "developer"))

    cc = CSVConf("hello.csv", init_columns=["C", "D"])
    cc.append({"A": 3, "B": 4}).append({"A": 3, "B": 4}).append({"A": 5, "B": 6}).save()
