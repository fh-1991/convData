import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

def profile_data_to_excel(df, output_file='data_profile_report.xlsx'):
    """データフレームの各カラムをプロファイリングし、Excelファイルに出力"""
    
    # 各データ型別の結果を格納するリスト
    numeric_results = []
    datetime_results = []
    categorical_results = []
    
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
        
        # 基本情報（全データ型共通）
        base_info = {
            'カラム名': col,
            'データ型': str(df[col].dtype),
            '総行数': len(df),
            '非null値数': len(data),
            'ユニーク値数': df[col].nunique(),
            '欠損値数': df[col].isnull().sum(),
            '欠損率(%)': round(df[col].isnull().mean() * 100, 2)
        }
        
        # 数値データの場合
        if pd.api.types.is_numeric_dtype(df[col]):
            print(f"平均: {data.mean():.2f}")
            print(f"標準偏差: {data.std():.2f}")
            print(f"最小値: {data.min()}")
            print(f"最大値: {data.max()}")
            print(f"中央値: {data.median():.2f}")
            
            # 数値データの統計情報を追加
            numeric_info = base_info.copy()
            numeric_info.update({
                '平均': round(data.mean(), 2) if len(data) > 0 else None,
                '標準偏差': round(data.std(), 2) if len(data) > 0 else None,
                '最小値': data.min() if len(data) > 0 else None,
                '最大値': data.max() if len(data) > 0 else None,
                '中央値': round(data.median(), 2) if len(data) > 0 else None,
                '25%分位': round(data.quantile(0.25), 2) if len(data) > 0 else None,
                '75%分位': round(data.quantile(0.75), 2) if len(data) > 0 else None
            })
            numeric_results.append(numeric_info)
            
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
            
            # 日付データの統計情報を追加
            datetime_info = base_info.copy()
            datetime_info.update({
                '最古日付': data.min() if len(data) > 0 else None,
                '最新日付': data.max() if len(data) > 0 else None,
                '期間': str(data.max() - data.min()) if len(data) > 0 else None,
                '年数': round((data.max() - data.min()).days / 365.25, 2) if len(data) > 0 else None
            })
            datetime_results.append(datetime_info)
            
            # 日付分布
            plt.hist(data, bins=30, alpha=0.7, edgecolor='black')
            plt.title(f'{col}\n(日付)')
            plt.xlabel(col)
            plt.ylabel('頻度')
            plt.xticks(rotation=45)
            
        # カテゴリ・文字列データの場合
        else:
            top_values = data.value_counts().head(10)
            mode_value = data.mode().iloc[0] if len(data.mode()) > 0 else 'なし'
            print(f"最頻値: {mode_value}")
            print(f"トップ5値:\n{top_values.head()}")
            
            # カテゴリデータの統計情報を追加
            categorical_info = base_info.copy()
            categorical_info.update({
                '最頻値': mode_value,
                '最頻値の出現回数': data.value_counts().iloc[0] if len(data) > 0 else 0,
                '最頻値の割合(%)': round(data.value_counts().iloc[0] / len(data) * 100, 2) if len(data) > 0 else 0,
                'トップ2番目': data.value_counts().index[1] if len(data.value_counts()) > 1 else None,
                'トップ3番目': data.value_counts().index[2] if len(data.value_counts()) > 2 else None
            })
            categorical_results.append(categorical_info)
            
            # 棒グラフ（上位10個まで）
            if len(top_values) > 10:
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
    
    # Excelファイルに出力
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 数値データシート
        if numeric_results:
            numeric_df = pd.DataFrame(numeric_results)
            numeric_df.to_excel(writer, sheet_name='数値データ', index=False)
            print(f"\n数値データ: {len(numeric_results)}カラムをExcelに出力")
        
        # 日付データシート
        if datetime_results:
            datetime_df = pd.DataFrame(datetime_results)
            datetime_df.to_excel(writer, sheet_name='日付データ', index=False)
            print(f"日付データ: {len(datetime_results)}カラムをExcelに出力")
        
        # カテゴリデータシート
        if categorical_results:
            categorical_df = pd.DataFrame(categorical_results)
            categorical_df.to_excel(writer, sheet_name='カテゴリデータ', index=False)
            print(f"カテゴリデータ: {len(categorical_results)}カラムをExcelに出力")
        
        # 全体概要シート
        summary_data = {
            '項目': ['総行数', '総列数', '数値カラム数', '日付カラム数', 'カテゴリカラム数'],
            '値': [len(df), len(df.columns), len(numeric_results), len(datetime_results), len(categorical_results)]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='全体概要', index=False)
        
        print(f"\nExcelファイル '{output_file}' を作成しました。")
        print("シート: '数値データ', '日付データ', 'カテゴリデータ', '全体概要'")

# 使用例
if __name__ == "__main__":
    # サンプルデータの作成
    np.random.seed(42)
    sample_data = {
        '年齢': np.random.normal(35, 10, 1000),
        '給与': np.random.lognormal(10, 0.5, 1000),
        '部署': np.random.choice(['営業', '技術', '管理', '企画'], 1000),
        '入社日': pd.date_range('2020-01-01', periods=1000, freq='D'),
        '評価': np.random.choice(['A', 'B', 'C', 'D'], 1000, p=[0.1, 0.3, 0.4, 0.2]),
        '売上': np.random.exponential(1000, 1000),
        '更新日': pd.date_range('2023-01-01', periods=1000, freq='H'),
        '地域': np.random.choice(['東京', '大阪', '名古屋', '福岡'], 1000)
    }
    
    df = pd.DataFrame(sample_data)
    
    # 一部欠損値を追加
    df.loc[np.random.choice(df.index, 50), '年齢'] = np.nan
    df.loc[np.random.choice(df.index, 30), '部署'] = np.nan
    
    # プロファイリング実行（Excel出力付き）
    profile_data_to_excel(df, 'データプロファイリング結果.xlsx')
    
    # CSVファイルから読み込む場合の例
    # df = pd.read_csv('your_file.csv')
    # profile_data_to_excel(df, 'your_profile_report.xlsx')
