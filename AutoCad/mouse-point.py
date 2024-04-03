from pynput.mouse import Listener


def on_click(x, y, button, pressed):
    if pressed:
        print(f"pyautogui.click({x}, {y})")


# 클릭 이벤트 핸들러 등록
with Listener(on_click=on_click) as listener:
    listener.join()
