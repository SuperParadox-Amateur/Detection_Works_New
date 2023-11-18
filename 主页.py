'''
主页
'''
# %% 导入库
import streamlit as st

# %% 设置
st.set_page_config(layout="wide", initial_sidebar_state="auto")

# %% 主页

# streamlit run main.py --server.port 8080

st.markdown('''# 脚本应用的使用说明

## 前提

### 拒绝加密的文件

''')

st.info("不要处理加密的文件！")

st.markdown('''要使用本页面下的所有功能，首要前提就是**文件不能被亿赛通加密**！

被亿赛通加密的文件无法被处理。

### 使用非IE内核的浏览器

为保证各项功能正常使用，不要使用IE内核的浏览器，要改用Chrome内核的浏览器（例如Google Chrome浏览器、Microsoft Edge浏览器、360极速浏览器和Vivaldi浏览器等等）或者Firefox浏览器。

其他国产双核浏览器也可以使用，但是必须在极速模式下才可以使用，因为兼容模式是使用IE内核。

## 使用方法

'''
)