import re
import os
import win32com.client as win32


def convert_latex_to_word_math(file_path):

  word = win32.gencache.EnsureDispatch("Word.Application")
  word.Visible = True  # 设置为 True 可以直观看到 Word 正在处理

  try:
    doc = word.Documents.Open(file_path)


    pattern = re.compile(r"(\$\$([\s\S]*?)\$\$|\$([^\$]+?)\$)")

    print("正在扫描并转换公式...")

    for para in doc.Paragraphs:
      para_range = para.Range
      text = para_range.Text

      matches = list(pattern.finditer(text))
      if not matches:
        continue

      for match in reversed(matches):
        start_idx = match.start()
        end_idx = match.end()
        full_match_text = match.group(1)

        if full_match_text.startswith("$$"):
          latex_code = match.group(2).strip()
        else:
          latex_code = match.group(3).strip()

        sub_range = doc.Range(
            para_range.Start + start_idx, para_range.Start + end_idx
        )


        sub_range.Text = latex_code


        doc.OMaths.Add(sub_range)


        try:
          sub_range.OMaths.Item(1).BuildUp()
        except Exception as e:
          print(f"公式渲染失败 [{latex_code}]: {e}")


    base_name, ext = os.path.splitext(file_path)
    output_path = f"{base_name}_转换后{ext}"
    # 另存为新文件
    doc.SaveAs(output_path)
    print(f"转换成功！已另存为新文件: {output_path}")


  except Exception as e:
    print(f"发生错误: {e}")
  finally:

    doc.Close(False)
    word.Quit()


if __name__ == "__main__":
  target_docx = r" YOUR PATH"
  convert_latex_to_word_math(target_docx)
