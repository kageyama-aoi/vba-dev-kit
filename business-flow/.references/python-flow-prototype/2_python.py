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
    steps = [
        ("class_select", step_class_select),
        ("switch_detail", step_switch_detail),
        ("apply", step_apply),
        ("fill_dates", step_fill_dates),
        ("course_set", step_course_set),
        ("transaction", step_transaction),
        ("verify", step_verify),
    ]

    # スキップ対象
    skip_steps = set()

    # error_injection = {
    #     "breakValue": "SKIP",
    #     "breakTarget": "course_set"
    # }

    # if error_injection["breakValue"] == "SKIP":
    #     skip_steps.add(error_injection["breakTarget"])
        
    error_injection = {
        "breakValue": "SKIP",
        "targets": ["switch_detail", "course_set"]
    }

    if error_injection["breakValue"] == "SKIP":
        skip_steps.update(error_injection["targets"])

    # 実行計画を生成
    plan = [(n, f) for (n, f) in steps if n not in skip_steps]

    print("=== 実行計画 ===")
    for name, _ in plan:
        print(name)

    print("=== 実行開始 ===")
    for _, func in plan:
        func()


if __name__ == "__main__":
    main()


