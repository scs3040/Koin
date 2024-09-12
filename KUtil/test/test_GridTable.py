import wx
import wx.grid
import wx.adv

class Example(wx.Frame):
    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, '그리드 예제', size=(600, 400))
        # 사용자 코드
        panel = wx.Panel(self, -1)


        # 그리드 클래스의 오브젝트 생성
        grid = wx.grid.Grid(panel, -1)

        # 그리드 행, 열 개수 설정 : 행 10개, 열 5개
        grid.CreateGrid(10, 5)

        # 셀 행 높이 설정 (픽셀크기)
        grid.SetRowSize(0, 60)
        grid.SetRowSize(1, 30)
        grid.SetRowSize(2, 60)

        # 열 너비 설정 (픽셀크기)
        grid.SetColSize(0, 120)
        grid.SetColSize(1, 100)
        grid.SetColSize(2, 80)

        # 셀(0,0) 값 설정
        grid.SetCellValue(0, 0, '0행 0열')
        grid.SetCellValue(1, 0, '1행 0열')
        grid.SetCellValue(2, 0, '2행 0열')

        # 읽기전용 셀 설정
        grid.SetCellValue(0, 1, '(읽기전용)')
        grid.SetReadOnly(0, 1)

        # 색상 설정
        grid.SetCellValue(1, 1, '색상설정')
        grid.SetCellTextColour(1,1, wx.RED)
        grid.SetCellBackgroundColour(1,1, (150,220,220))

        # 컬럼 포멧 설정
        grid.SetColFormatFloat(2, 6, 5) # (컬럼번호, 문자열 길이(width), 정밀도(precision))
        grid.SetCellValue(0, 2, '3.1415926535')
        grid.SetCellValue(1, 2, '1.2345678901')
        grid.SetCellValue(2, 2, '1.4285714285')


        # 그리드를 담을 박스사이저 생성
        bsizer = wx.BoxSizer(wx.VERTICAL)
        bsizer.Add(grid, -1, wx.ALL, 20)
        panel.SetSizer(bsizer)

        self.Show()



if __name__ == '__main__':

    app = wx.App()
    frame = Example(parent=None, id=-1)
    app.MainLoop()