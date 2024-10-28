import typing


def format_duration(time_duration: typing.Union[int, float], only_sec=False) -> str:
    """get time duration of the given time in secs
    """
    # 如果只需要显示秒，则直接返回
    if only_sec:
        return str(round(time_duration, 2)) + "s"

    time_duration = int(time_duration)
    units = ["s", "min", "h"]
    max_level = 2  # if level = 3 will exceed the max level

    expression_list = []

    def get_format_of(duration):
        nonlocal units, max_level, expression_list

        cur_level = -1
        cur_size = duration
        while True:
            cur_level += 1
            res_without_unit, rest_size = divmod(cur_size, 60)
            expression_list.append(f"{rest_size}{units[cur_level]}")
            if res_without_unit == 0:
                break
            cur_size = res_without_unit
        # max is 1TB
        if cur_level > max_level:
            return "more than 60 h"
        return ":".join(expression_list[::-1])

    return get_format_of(time_duration)


if __name__ == '__main__':
    print(format_duration(1.12321321321, only_sec=True))
