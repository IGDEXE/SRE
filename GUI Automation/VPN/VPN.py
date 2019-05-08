# Configurar FortiClient
# Importa as blibliotecas
import pyautogui
import time
import sys

# Recebe os valores
user = sys.argv[1]
pws = sys.argv[2]

# Abre o executar do Windows e busca a pasta
pyautogui.hotkey('win','r');pyautogui.typewrite('C:\\Program Files\\Fortinet\\FortiClient');pyautogui.press('enter')
time.sleep(2)
pyautogui.hotkey('alt','space','x')
# Abre o arquivo
pyautogui.typewrite('FortiClient');pyautogui.press('enter')
time.sleep(4)
# Abre a tela de acesso remoto
pyautogui.click(359,399)
time.sleep(2)
# Configura o login
pyautogui.doubleClick(870,511)
pyautogui.click(870,511)
pyautogui.typewrite(user)
# Configura a senha
pyautogui.doubleClick(870,545)
pyautogui.typewrite(pws)
# Faz o login
pyautogui.click(829,620)
# Aceita o certificado
time.sleep(5)
pyautogui.click(1041,100)
pyautogui.click(612,487)
