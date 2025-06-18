import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime

def profile_data(df):
    """データフレームの各カラムをプロファイリングし、ヒストグラムを作成"""
    
    # 基本情報の表示
    print("=== データ概要 ===")
    print(f"行数: {len(df)}")
    print(f"列数: {len(df.columns)}")
    print(f"データ型:\n{df.dtypes}")
    print(f"\n欠損値:\n{df.isnull().sum()}")
    
    # カラム数に応じてプロット配置を決定
    n_cols = len(df.columns)
    fig_cols = min(3, n_cols)
    fig_rows = (n_cols + fig_cols - 1) // fig_cols
    
    plt.figure(figsize=(15, 5 * fig_rows))
    
    for i, col in enumerate(df.columns):
        plt.subplot(fig_rows, fig_cols, i + 1)
        
        # 欠損値を除外
        data = df[col].dropna()
        
        print(f"\n=== {col} ===")
        print(f"データ型: {df[col].dtype}")
        print(f"非null値数: {len(data)}")
        print(f"ユニーク値数: {df[col].nunique()}")
        print(f"欠損率: {df[col].isnull().mean():.2%}")
        
        # 数値データの場合
        if pd.api.types.is_numeric_dtype(df[col]):
            print(f"平均: {data.mean():.2f}")
            print(f"標準偏差: {data.std():.2f}")
            print(f"最小値: {data.min()}")
            print(f"最大値: {data.max()}")
            print(f"中央値: {data.median():.2f}")
            
            # ヒストグラム
            plt.hist(data, bins=30, alpha=0.7, edgecolor='black')
            plt.title(f'{col}\n(数値)')
            plt.xlabel(col)
            plt.ylabel('頻度')
            
        # 日付・時刻データの場合
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            print(f"最古: {data.min()}")
            print(f"最新: {data.max()}")
            print(f"期間: {data.max() - data.min()}")
            
            # 日付分布
            plt.hist(data, bins=30, alpha=0.7, edgecolor='black')
            plt.title(f'{col}\n(日付)')
            plt.xlabel(col)
            plt.ylabel('頻度')
            plt.xticks(rotation=45)
            
        # カテゴリ・文字列データの場合
        else:
            top_values = data.value_counts().head(10)
            print(f"最頻値: {data.mode().iloc[0] if len(data.mode()) > 0 else 'なし'}")
            print(f"トップ5値:\n{top_values.head()}")
            
            # 棒グラフ（上位10個まで）
            if len(top_values) > 10:
                # 多すぎる場合は上位10個のみ
                plot_data = top_values.head(10)
                title_suffix = f"(上位10/{df[col].nunique()})"
            else:
                plot_data = top_values
                title_suffix = ""
                
            plt.bar(range(len(plot_data)), plot_data.values, alpha=0.7)
            plt.title(f'{col}\n(カテゴリ){title_suffix}')
            plt.xlabel('カテゴリ')
            plt.ylabel('頻度')
            plt.xticks(range(len(plot_data)), plot_data.index, rotation=45, ha='right')
    
    plt.tight_layout()
    plt.show()

# 使用例
if __name__ == "__main__":
    # サンプルデータの作成
    np.random.seed(42)
    sample_data = {
        '年齢': np.random.normal(35, 10, 1000),
        '給与': np.random.lognormal(10, 0.5, 1000),
        '部署': np.random.choice(['営業', '技術', '管理', '企画'], 1000),
        '入社日': pd.date_range('2020-01-01', periods=1000, freq='D'),
        '評価': np.random.choice(['A', 'B', 'C', 'D'], 1000, p=[0.1, 0.3, 0.4, 0.2])
    }
    
    df = pd.DataFrame(sample_data)
    
    # 一部欠損値を追加
    df.loc[np.random.choice(df.index, 50), '年齢'] = np.nan
    df.loc[np.random.choice(df.index, 30), '部署'] = np.nan
    
    # プロファイリング実行
    profile_data(df)
    
    # CSVファイルから読み込む場合の例
    # df = pd.read_csv('your_file.csv')
    # profile_data(df)
