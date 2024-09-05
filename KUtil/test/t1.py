import tkinter as tk
from tkinter import filedialog

def get_filename():
    filename = filedialog.askopenfilename(title="파일 선택", filetypes=(("모든 파일", "*.*"),))
    if filename:
        label.config(text=f"선택한 파일명: {filename}")
    else:
        label.config(text="파일을 선택하지 않았습니다.")

# tkinter 윈도우 생성
root = tk.Tk()
root.title("파일명 입력 대화상자")

# 파일명을 표시할 라벨
label = tk.Label(root, text="파일명을 선택하세요.", padx=10, pady=10)
label.pack()

# 파일 선택 버튼
button = tk.Button(root, text="파일 선택", command=get_filename)
button.pack()

# 윈도우 실행
root.mainloop()