
if __name__ == '__main__':
    from yrx_project.init_env import init_env, run_app
    # 环境初始化
    init_env("prod")

    # 启动程序
    run_app(1400, 1000)
