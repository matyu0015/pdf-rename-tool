"""
テスト用の学生写真ファイルを作成するスクリプト
学生IDをファイル名とした画像を生成
"""

from PIL import Image, ImageDraw, ImageFont
import os

def create_student_photos():
    # 写真フォルダを作成
    photo_dir = "test_student_photos"
    if not os.path.exists(photo_dir):
        os.makedirs(photo_dir)

    # テスト用の学生ID一覧（エクセルと同じ）
    student_ids = [
        "20240001", "20240002", "20240003", "20240004", "20240005",
        "20240006", "20240007", "20240008", "20240009", "20240010"
    ]

    # 学生名（顔写真の代わりに名前を表示）
    student_names = [
        "山田太郎", "佐藤花子", "鈴木一郎", "田中美咲", "高橋健太",
        "伊藤優子", "渡辺翔太", "中村さくら", "小林大輔", "加藤愛美"
    ]

    # 色のバリエーション
    colors = [
        "#FF6B6B", "#4ECDC4", "#45B7D1", "#FFA07A", "#98D8C8",
        "#F7DC6F", "#BB8FCE", "#85C1E2", "#F8B195", "#A8E6CF"
    ]

    print(f"学生写真ファイルを作成中...")

    for idx, (student_id, name) in enumerate(zip(student_ids, student_names)):
        # 画像サイズ 200x250
        img = Image.new('RGB', (200, 250), color=colors[idx])
        draw = ImageDraw.Draw(img)

        # システムフォントを使用（日本語対応）
        try:
            # macOS/Linuxの場合
            font_large = ImageFont.truetype("/System/Library/Fonts/ヒラギノ角ゴシック W4.ttc", 24)
            font_small = ImageFont.truetype("/System/Library/Fonts/ヒラギノ角ゴシック W4.ttc", 18)
        except:
            try:
                # Windowsの場合
                font_large = ImageFont.truetype("C:\\Windows\\Fonts\\msgothic.ttc", 24)
                font_small = ImageFont.truetype("C:\\Windows\\Fonts\\msgothic.ttc", 18)
            except:
                # フォントが見つからない場合はデフォルト
                font_large = ImageFont.load_default()
                font_small = ImageFont.load_default()

        # 学生IDを描画（上部）
        text_id = student_id
        bbox = draw.textbbox((0, 0), text_id, font=font_small)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        position_id = ((200 - text_width) // 2, 40)
        draw.text(position_id, text_id, fill='white', font=font_small)

        # 学生名を描画（中央）
        text_name = name
        bbox = draw.textbbox((0, 0), text_name, font=font_large)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        position_name = ((200 - text_width) // 2, 110)
        draw.text(position_name, text_name, fill='white', font=font_large)

        # "Student Photo"を描画（下部）
        text_label = "Student Photo"
        bbox = draw.textbbox((0, 0), text_label, font=font_small)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        position_label = ((200 - text_width) // 2, 200)
        draw.text(position_label, text_label, fill='white', font=font_small)

        # ファイルを保存（JPG形式）
        output_path = os.path.join(photo_dir, f"{student_id}.jpg")
        img.save(output_path, "JPEG", quality=95)

    print(f"✓ {len(student_ids)}枚の学生写真ファイルを作成しました: {photo_dir}/")
    print(f"  - ファイル形式: JPG")
    print(f"  - ファイル名: {student_ids[0]}.jpg 〜 {student_ids[-1]}.jpg")

    # 追加で2枚、学生ID以外の画像も作成（マッチしないケースのテスト用）
    extra_ids = ["99999999", "88888888"]
    extra_names = ["テスト用A", "テスト用B"]

    for student_id, name in zip(extra_ids, extra_names):
        img = Image.new('RGB', (200, 250), color='#999999')
        draw = ImageDraw.Draw(img)

        try:
            font_large = ImageFont.truetype("/System/Library/Fonts/ヒラギノ角ゴシック W4.ttc", 24)
            font_small = ImageFont.truetype("/System/Library/Fonts/ヒラギノ角ゴシック W4.ttc", 18)
        except:
            try:
                font_large = ImageFont.truetype("C:\\Windows\\Fonts\\msgothic.ttc", 24)
                font_small = ImageFont.truetype("C:\\Windows\\Fonts\\msgothic.ttc", 18)
            except:
                font_large = ImageFont.load_default()
                font_small = ImageFont.load_default()

        text_id = student_id
        bbox = draw.textbbox((0, 0), text_id, font=font_small)
        text_width = bbox[2] - bbox[0]
        position_id = ((200 - text_width) // 2, 40)
        draw.text(position_id, text_id, fill='white', font=font_small)

        text_name = name
        bbox = draw.textbbox((0, 0), text_name, font=font_large)
        text_width = bbox[2] - bbox[0]
        position_name = ((200 - text_width) // 2, 110)
        draw.text(position_name, text_name, fill='white', font=font_large)

        output_path = os.path.join(photo_dir, f"{student_id}.jpg")
        img.save(output_path, "JPEG", quality=95)

    print(f"  + テスト用の追加画像: {len(extra_ids)}枚（エクセルにマッチしないID）")

if __name__ == "__main__":
    create_student_photos()
