import os

import pandas as pd


def clean_access_data():
        file_path = os.path.join(os.getcwd(), "Book1.xlsx")
        builder_path = builder_data="BuilderList.xlsx"  # Assuming it's also in CSV folder

        # 🧾 Load Excel data
        access_df = pd.read_excel(file_path)
        builder_df = pd.read_excel(builder_path)

        # 🏷️ Get the list of valid 得意先名 from builder file
        valid_names = builder_df['builder'].dropna().unique()

        # 🔍 Match and separate the data
        matched_df = access_df[access_df['得意先名'].isin(valid_names)]
        unmatched_df = access_df[~access_df['得意先名'].isin(valid_names)]

        # 🖨️ Print matched and unmatched values
        print("✅ 以下の得意先名はBuilderListに存在しているため残されました:")
        print(matched_df['得意先名'].unique())

        print("\n❌ 以下の得意先名はBuilderListに存在していないため削除されました:")
        print(unmatched_df['得意先名'].unique())

        # 💾 Save matched records back to the original file
        matched_df.to_excel(file_path, index=False)

        print(f"\nファイルが保存されました: {file_path}")
clean_access_data()