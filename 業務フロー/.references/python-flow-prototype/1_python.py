def step_class_select():
    print("① クラス選択")

def step_switch_detail():
    print("② 詳細画面へ")

def step_apply():
    print("③ クラス適用")

def step_fill_dates():
    print("④ 日付入力")

def step_course_set():
    print("⑤ コース設定")

def step_transaction():
    print("⑥ トランザクション")

def step_verify():
    print("⑦ エラーチェック")


def main():
    # 行程表（実行計画）
    steps = [
        ("class_select", step_class_select),
        ("switch_detail", step_switch_detail),
        ("apply", step_apply),
        ("fill_dates", step_fill_dates),
        ("course_set", step_course_set),
        ("transaction", step_transaction),
        ("verify", step_verify),
    ]

    for name, func in steps:
        func()


if __name__ == "__main__":
    main()