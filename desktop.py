from random import randint
from tkinter import *


class Randball():
    def __init__(self, canvas, scrnwidth, scrnheight):
        # 初始化画布
        self.canvas = canvas
        # 初始化球的圆心坐标
        self.x_pos = randint(50, int(scrnwidth))  # X轴的坐标 randint 随机产生一个范围内的整数
        self.y_pos = randint(50, int(scrnheight))  # Y轴的坐标
        # 球的移动距离
        self.x_move = 10
        self.y_move = 10
        # 整个屏幕的高和宽
        self.scrnwidth = scrnwidth
        self.scrnheight = scrnheight
        # 初始化球的半径
        self.randius = randint(10, 80)
        # 随机产生球的颜色
        rc = lambda: randint(0, 255)
        self.color = '#%02x%02x%02x' % (rc(), rc(), rc())

    def create_ball(self):
        # 计算得到用于创建球的四个坐标
        x1 = self.x_pos - self.randius
        y1 = self.y_pos - self.randius
        x2 = self.x_pos + self.randius
        y2 = self.y_pos + self.randius
        # 画球
        self.item = self.canvas.create_oval(x1, y1, x2, y2, fill=self.color, outline=self.color)

    def move_ball(self):
        # 球按照指定距离移动，如果碰到障碍就向相反的方向运动
        self.x_pos += self.x_move
        self.y_pos += self.y_move
        if self.x_pos >= self.scrnwidth - self.randius:
            self.x_move = -self.x_move
        elif self.y_pos >= self.scrnheight - self.randius:
            self.y_move = -self.y_move
        elif self.x_pos < self.randius:
            self.x_move = abs(self.x_move)
        elif self.y_pos < self.randius:
            self.y_move = abs(self.y_move)
        self.canvas.move(self.item, self.x_move, self.y_move)


class Screensaver():
    balls = []

    def __init__(self, ball_nums):
        self.win = Tk()
        self.width = self.win.winfo_screenwidth()
        self.height = self.win.winfo_screenheight()
        self.win.overrideredirect(True)
        self.win.attributes('-alpha', 1)
        # 绑定事件，有任何动作退出屏保
        self.win.bind('<Any-Button>', self.exit_screensaver)
        self.win.bind('<Motion>', self.exit_screensaver)
        self.canvas = Canvas(self.win, width=self.width, height=self.height, bg="#FFFFFF")  # 背景 颜色自己随便调整，至于啥颜色就看自己的心情了
        self.canvas.pack()

        for i in range(0, ball_nums):
            ball = Randball(self.canvas, scrnwidth=self.width, scrnheight=self.height)
            ball.create_ball()
            self.balls.append(ball)
        self.run_screensaver()
        self.win.mainloop()

    def run_screensaver(self):
        for ball in self.balls:
            ball.move_ball()
        self.canvas.after(30, self.run_screensaver)

    def exit_screensaver(self, event):
        self.win.destroy()


def main():
    Screensaver(30)  # 屏保上球的个数


if __name__ == '__main__':
    main()
