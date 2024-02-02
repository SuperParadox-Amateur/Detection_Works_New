'''
选项词典
'''

from typing import Dict

# 字体
_fonts_dict: Dict["str", "str"] = {
    "雅黑": "C:/Windows/Fonts/msyh.ttf",
    "宋体": "C:/Windows/Fonts/simsun.ttc",
    "仿宋": "C:/Windows/Fonts/simfang.ttf",
    "黑体": "C:/Windows/Fonts/simhei.ttf",
    "楷体": "C:/Windows/Fonts/simkai.ttf"
}

# 方向
_directions_dict: Dict["str", "str"] = {
    "从左到右": "ltr",
    "从右到左": "rtl",
    "从上到下": "ttb"
}

# 对齐
_align_dict: Dict["str", "str"] = {
    "靠左": "left",
    "居中": "center",
    "靠右": "right"
}

# with open(r"info_files\腾讯文档信息.txt", "r", encoding="utf-8") as f:
#     cookie_data: str = f.read()


# tx_doc_info_dict: dict = {
    # excel文档地址
    # "document_url": 'https://docs.qq.com/sheet/DZW5mWlBSeXpPb29N',
    # 此值每一份腾讯文档有一个,需要手动获取
    # "local_pad_id": '300000000$enfZPRyzOooM',
    # 打开腾讯文档后,从抓到的接口中获取cookie信息
#     "cookie_value": cookie_data
# }
