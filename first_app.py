# -*- coding: utf-8 -*-
"""
Created on Fri May 28 11:13:50 2021

@author: Jake
"""
import sys
from streamlit import cli as stcli
import streamlit as st
import numpy as np
import pandas as pd



# def main():
st.title('My first app')

# if __name__ == '__main__':
#     if st._is_running_with_streamlit:
#         main()
#     else:
#         sys.argv = ["streamlit", "run", sys.argv[0]]
#         sys.exit(stcli.main())