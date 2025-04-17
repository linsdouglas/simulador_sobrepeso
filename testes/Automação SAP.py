#Automação SAP
import time
import math
import  datetime
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinter import ttk
import customtkinter as ctk
from PIL import Image
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
import pyautogui as pt
import yagmail
from openpyxl import load_workbook
import comtypes.client
import glob
import shutil
import subprocess
import pandas as pd
import os
import sys