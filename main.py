# main.py
import os
from generator import generate_excel

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    json_dir = os.path.join(base_dir, "input_jsons")
    template_path = os.path.join(base_dir, "成品.xlsx")
    skills_path = os.path.join(base_dir, "《奇点之后》车卡序列资料正式版V0.6.6).xlsx")
    output_path = os.path.join(base_dir, "成品输出.xlsx")

    json_files = [
        os.path.join(json_dir, f)
        for f in os.listdir(json_dir)
        if f.endswith(".json")
    ]

    if not json_files:
        print("⚠️ 没有找到 JSON 文件")
        return

    print(f"检测到 {len(json_files)} 个角色，将生成 Excel...")
    generate_excel(template_path, json_files, skills_path, output_path)

if __name__ == "__main__":
    main()
