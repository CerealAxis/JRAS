# coding='utf-8'
import argparse
import os
import dashscope
import pypandoc
from http import HTTPStatus
from urllib.parse import quote

dashscope.api_key = ""

# 1. 定义生成商品链接的函数
def generate_good_link(product_id):
    """
    根据给定的商品ID，生成京东商品链接。

    参数:
    - product_id (int): 商品ID

    返回:
    str: 生成的商品链接字符串
    """
    base_url = "https://item.jd.com/{pid}.html"
    return base_url.format(pid=product_id)


# 2. 读取TXT文件内容的函数
def read_txt_file(file_path):
    """
    读取指定路径下的TXT文件内容。如果读取过程中发生错误，打印错误信息并返回None。

    参数:
    - file_path (str): TXT文件路径

    返回:
    str or None: 成功读取则返回文件内容，否则返回None
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            txt_content = f.read()
        return txt_content
    except Exception as e:
        print(f"读取TXT文件时发生错误: {e}")
        return None


# 3. Markdown转Docx的函数
def markdown_to_docx(content, output_path):
    """
    将Markdown格式的内容转换为Docx文档，并保存到指定路径。转换过程中可能发生的错误会被捕获并打印。

    参数:
    - content (str): Markdown格式的内容
    - output_path (str): 目标Docx文档的保存路径
    """
    # 将Markdown内容暂存到临时文件
    with open('temp.md', 'w', encoding='utf-8') as f:
        f.write(content)

    try:
        converted_text = pypandoc.convert_file('temp.md', 'docx', outputfile=output_path)
    except (IOError, OSError, RuntimeError) as e:
        print(f"转换过程中发生错误: {e}")
    else:
        print("------ 分析报告已生成并保存至{} ------".format(output_path))

    # 删除临时文件
    os.remove('temp.md')


# 4. 构建系统消息的函数
def build_system_message(product_id):
    """
    根据商品ID构建一条系统消息，包含商品链接和提示用户撰写分析报告的文本。

    参数:
    - product_id (int): 商品ID

    返回:
    dict: 系统消息字典
    """
    good_link = generate_good_link(product_id)
    return {
        'role': 'system',
        'content': f'针对京东商品（{quote(good_link)}），我将提供商品链接及用户评价摘要，基于这些信息，请撰写一份详实、全面的商品分析报告，并给出具有针对性的改进建议。字数不少于2千字，格式要书面、美观，并尽可能地避免使用过于简单的语言。'
    }


# 5. 主要调用逻辑的函数
def call_with_messages(product_id):
    """
    根据商品ID，执行一系列操作以生成商品分析报告并保存为Docx文档。

    参数:
    - product_id (int): 商品ID
    """
    txt_summary = read_txt_file("output/contents.txt")
    good_link = generate_good_link(product_id)

    system_message = build_system_message(product_id)
    user_message = {
        'role': 'user',
        'content': [{'text': f"{good_link}{txt_summary}"}]
    }

    messages = [system_message, user_message]

    response = dashscope.Generation.call(
        dashscope.Generation.Models.qwen_turbo,
        messages=messages,
        result_format='message',
    )

    if response.status_code == HTTPStatus.OK:
        recipe = response['output']['choices'][0]['message']['content']

        # 将Markdown格式的结果转换为Word文档
        markdown_to_docx(recipe, "output/report.docx")
    else:
        print('请求失败：',
              f"Request id: {response.request_id}, ",
              f"Status code: {response.status_code}, ",
              f"Error code: {response.code}, ",
              f"Error message: {response.message}")


# 6. 主函数
def main():
    """
    解析命令行参数，获取商品ID，并调用call_with_messages()函数生成分析报告。
    """
    parser = argparse.ArgumentParser(description="生成分析报告并保存为Word文档")
    parser.add_argument('-p', '--product-id', required=True, type=int,
                        help="输入商品ID")

    args = parser.parse_args()
    product_id = args.product_id
    call_with_messages(product_id)


if __name__ == '__main__':
    main()
