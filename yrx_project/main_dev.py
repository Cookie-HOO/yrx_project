
if __name__ == '__main__':
    from yrx_project.init_env import init_env, run_app
    # 环境初始化
    init_env("dev")

    # 启动程序
    import openpyxl  # 显式引入，pandas读取excel需要
    run_app(1400, 1000)
