# Importa as blibliotecas
import pyautogui
import time

# Abre o executar do Windows e busca o servidor
pyautogui.hotkey('win','r');pyautogui.typewrite('\\\srsdcsp01\\NETLOGON\\Bitdefender\\Setup');pyautogui.press('enter')
time.sleep(2)
# Abre o arquivo de instalação
pyautogui.click(520,479);pyautogui.typewrite('epskit_x64');pyautogui.press('enter')
# Mostra mensagem na tela
pyautogui.alert(title='INFRA - BitDefender', text='Aguarde a instalacao, pode demorar um pouco', button='OK')
