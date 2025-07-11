def convert_hex_to_decimal(hex_string):
    """
    16進数文字列を指定されたルールで10進数文字列に変換する
    
    Args:
        hex_string (str): "0x"で始まる16進数文字列
    
    Returns:
        str: 変換された10進数文字列
    """
    # "0x"を除去
    if hex_string.startswith("0x"):
        hex_part = hex_string[2:]
    else:
        hex_part = hex_string
    
    # 8桁の場合は末尾2桁を除外
    if len(hex_part) == 8:
        hex_part = hex_part[:-2]
    
    # 2桁ずつに分割
    hex_pairs = []
    for i in range(0, len(hex_part), 2):
        pair = hex_part[i:i+2]
        hex_pairs.append(pair)
    
    # 各2桁の16進数を10進数に変換
    decimal_parts = []
    for hex_pair in hex_pairs:
        decimal_value = int(hex_pair, 16)
        # 1桁の場合は2桁に0埋め、2桁以上はそのまま
        if decimal_value < 10:
            decimal_parts.append(f"{decimal_value:02d}")
        else:
            decimal_parts.append(str(decimal_value))
    
    # 結果を連結
    return "".join(decimal_parts)


def main():
    # テストケース
    test_cases = [
    ]
    
    print("16進数から10進数への変換結果:")
    print("-" * 40)
    
    for hex_str in test_cases:
        result = convert_hex_to_decimal(hex_str)
        print(f"{hex_str} → {result}")
    
    print("-" * 40)
    
    # 対話的な入力
    while True:
        user_input = input("\n16進数文字列を入力してください (終了するには 'q' を入力): ").strip()
        if user_input.lower() == 'q':
            break
        
        try:
            result = convert_hex_to_decimal(user_input)
            print(f"変換結果: {result}")
        except ValueError as e:
            print(f"エラー: 無効な16進数文字列です - {e}")
        except Exception as e:
            print(f"エラー: {e}")


if __name__ == "__main__":
    main()
