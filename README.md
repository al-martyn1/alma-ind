# Индикатор присутствия на рабочем месте

Устройство - лампочка под управлением Ардуино (для прототипа).

Представляет собой штангу, устанавливаемую вертикально, на панели стола или межсотовой перегородке между кубиклами,
незначительно возвышаясь над ними. Позволяет визуально оценить статус сотрудника.

Из нижней части штанги выходит провод USB для сопряжения с компьютером (USB-UART).
Возможен вариант установки разъёма USB-C для использования кабеля нужной длины.

Также в нижней части штанги располагаются разъёмы:
1. Разъём пульта управления
2. Опциональный разъём питания

Режимы работы (статусы сотрудника):

- Красный - занят/busy
- Зелёный - свободен/available
- Желтый - нет на месте, отошел/away
- Выключен - не на работе/not at work


## Связь с компьютером

Связь с компьютером осуществляется по UART.

Протокол простой, NMEA-подобный:

```
$PALMA,IND,C,N[*XX] CR LF
!PALMA,IND,C,N[*XX] CR LF
```

Где 'C' - команда:

- С - command/control - команда от компьютера к устройству
- A - ack/acknowledgment - уведомление от устройства об обработке команды компьютера
- M - (initiative) message - сообщение, генерируемое устройством при изменении его состояния с пульта управления.

а N - режим работы:

- 0 - not at work
- 1 - available
- 2 - away
- 3 - busy


## Устройство

Устройство имеет в своём составе контроллер, отвечающий за логику его работы. 
Контроллер располагается в верхней части штанги под лампой индикации 
(возможно и расположение в нижней части штанги в добавляющемся в этом случае блоке управления,
или в блоке с USB-хабом).

В штанге предусмотрено пять сквозных отверстий для саморезов крепления к основе 
(вертикальная панель стола или стенка кубикла), или может крепится к основе на двусторонний скотч.

Устройство имеет выносной пульт с двумя кнопками. Выносной пульт опционален и подключается
через 3.5 стерео jerk в нижней части штанги.

Можно сделать версию устройства со встроенным USB-хабом в нижней части штанги
(в этом случае добавляется разъём питания хаба в нижней части штанги).


## Пульт управления

Выносной пульт управления имеет две апаратные кнопки:

1. Кнопка смены режима работы
2. Кнопка заверения работы

При воздействиях на обе кнопки устройство отсылает соответствующие сообщения на ПК (по UART).


### Кнопка смены режима работы

Кнопка смены режима работы производит переключение режима занят/свободен.

Если был свободен - переключаемся в режим занят/busy.
Если был любой другой режим - переключаемся в режим свободен/avaialble.

Кнопка должна быть легко доступна и нажимаема, желательно возвышение над корпусом пульта.

При смене режима работы устройство выдаёт сообщения:

```
$PALMA,IND,M,1[*XX] CR LF
```
или
```
$PALMA,IND,M,3[*XX] CR LF
```


### Кнопка завершения работы

Данная кнопка должна быть защищена от случайного нажатия: например, высокими бортами корпуса пульта.

При нажатии кнопки устройство переходит в режим 'not at work' и выдаёт сообщение:

```
$PALMA,IND,M,0[*XX] CR LF
```

## Приложение для ПК

Приложение располагается в System Tray и иконкой информирует о статусе.

Приложение позволяет менять статус через UI (ограниченно - только свободен/занят).

Приложение слушает событие LockWorkStation, переключает устройство в режим away, или, при разблокировке, в режим available.
При запуске - включает режим available.

Приложение слушает COM-порт, и реагирует на сообщения устройства.

Приложение может производить запуск внешних программ при смене статуса (ушел с работы - запускается скрипт регистрации рабочего времени в БИТРИКС).

Также приложение может отвечать (или рассылать) UDP или SNMP сообщения/ответы, которые можно использовать
для сбора по сети информации о статусе сотрудников в системах учёта рабочего времени.


## Заметки по реализации приложения для ПК

### Отслеживание блокировки в Windows

Да, в Windows можно программно отследить момент блокировки и разблокировки компьютера. Самый надежный и документированный способ — использование функций Windows API. [1, 2] 
Основные механизмы отслеживания

1. WTSRegisterSessionNotification: Это основная функция API для подписки на уведомления об изменении состояния сессии.
2. WM_WTSSESSION_CHANGE: Сообщение, которое Windows отправляет вашему окну после регистрации.
3. Параметры сообщения (wParam):
  - WTS_SESSION_LOCK (0x7) — компьютер заблокирован.
  - WTS_SESSION_UNLOCK (0x8) — компьютер разблокирован. [2, 3, 4, 5, 6, 7] 
   
Способы реализации1. Для приложений с графическим интерфейсом (C#, C++, Delphi)
Вам нужно зарегистрировать дескриптор окна (HWND) для получения уведомлений.

* Регистрация: Вызовите WTSRegisterSessionNotification(hWnd, NOTIFY_FOR_THIS_SESSION).
* Обработка: В цикле обработки сообщений (WndProc) ловите WM_WTSSESSION_CHANGE и проверяйте значение wParam.
* Снятие регистрации: При закрытии программы вызовите WTSUnRegisterSessionNotification(hWnd). [3, 7, 8, 9] 

#### 2. Для .NET (C# / VB.NET)
Можно использовать класс Microsoft.Win32.SystemEvents, который упрощает работу с API:

* Подпишитесь на событие SessionSwitch.
* Проверяйте e.Reason == SessionSwitchReason.SessionLock. [8] 

#### 3. Для консольных приложений и сервисов
Поскольку у консольных программ нет окна по умолчанию, вам придется:

* Создать скрытое («невидимое») окно специально для приема сообщений.
* Для системных сервисов (Windows Services) учитывайте «изоляцию нулевой сессии» (Session 0 isolation): сервис должен запускаться с правами взаимодействия с рабочим столом или использовать специальные механизмы для отслеживания сессий пользователей. [1, 10] 

#### Важные нюансы

* Права доступа: Обычно специальных административных прав для этого не требуется, если вы отслеживаете только свою сессию (NOTIFY_FOR_THIS_SESSION).
* Event ID: Если вам нужно анализировать блокировки постфактум, Windows записывает их в журнал событий под ID 4800 (заблокировано) и ID 4801 (разблокировано).
* Блокировка vs Скринсейвер: Блокировка и запуск экранной заставки — это разные события. Для заставки используются события 4802 и 4803. [3, 11, 12, 13] 

#### Ссылки

- [1] [https://stackoverflow.com](https://stackoverflow.com/questions/8606300/how-to-detect-windows-is-locked)
- [2] [https://francescofoti.com](https://francescofoti.com/2019/12/how-to-detect-windows-session-lock-unlock-in-vba/)
- [3] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/windows/win32/api/wtsapi32/nf-wtsapi32-wtsregistersessionnotification)
- [4] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/windows/win32/termserv/wm-wtssession-change#:~:text=Remarks.%20This%20message%20is%20sent%20only%20to,be%20painted%2C%20or%20triggering%20console%20animation%20effects.)
- [5] [https://learn.microsoft.com](https://learn.microsoft.com/nb-no/windows/win32/termserv/wm-wtssession-change#:~:text=Remarks.%20This%20message%20is%20sent%20only%20to,be%20painted%2C%20or%20triggering%20console%20animation%20effects.)
- [6] [https://www.bryancook.net](http://www.bryancook.net/2008/10/detect-wpf-lock-workstation-session.html)
- [7] [https://devblogs.microsoft.com](https://devblogs.microsoft.com/oldnewthing/20060104-50/?p=32783)
- [8] [https://www.bryancook.net](http://www.bryancook.net/2008/10/detect-wpf-lock-workstation-session.html)
- [9] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/windows/win32/api/wtsapi32/nf-wtsapi32-wtsunregistersessionnotification)
- [10] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/5758754/how-to-make-a-windows-service-detect-switch-user-e)
- [11] [https://www.manageengine.com](https://www.manageengine.com/products/active-directory-audit/kb/logon-logoff-events/event-id-4800.html#:~:text=Event%20ID%204800%20%E2%80%93%20The%20Workstaion%20Was,unlocks%20the%20workstation%2C%20event%204801%20is%20triggered.)
- [12] [https://superuser.com](https://superuser.com/questions/993178/windows-7-what-is-the-event-id-for-a-lock-event-and-how-to-tell-if-it-is-user-i#:~:text=When%20either%20a%20user%20manually%20locks%20his,an%20explanation%20of%20the%20sequence%20of%20events.)
- [13] [https://www.manageengine.com](https://www.manageengine.com/products/active-directory-audit/kb/logon-logoff-events/event-id-4802.html)


### Программная блокировка в Windows

Да, заблокировать компьютер программно можно. Самый простой и правильный способ в Windows — использовать системную функцию LockWorkStation из библиотеки user32.dll. [1, 2] 
Ниже приведены способы реализации на разных языках и через командную строку.

#### 1. Командная строка (CMD или PowerShell)
Вы можете выполнить блокировку без написания полноценного кода, используя утилиту rundll32: [3, 4] 

```
rundll32.exe user32.dll,LockWorkStation
```

Эту команду можно сохранить в .bat файл или создать для неё ярлык на рабочем столе. [5] 

#### 2. Python
В Python это делается через библиотеку ctypes, которая позволяет вызывать функции Windows API напрямую: [6, 7] 

import ctypes
ctypes.windll.user32.LockWorkStation()

#### 3. C# (.NET)
Для использования в C# необходимо импортировать функцию через DllImport: [8] 

```
using System.Runtime.InteropServices;

[DllImport("user32.dll")]public static extern bool LockWorkStation();
// Вызов функции
LockWorkStation();
```

#### 4. C++ (WinAPI)
В C++ достаточно подключить заголовочный файл windows.h и вызвать функцию напрямую: [9] 

```
#include <windows.h>
int main() {
    LockWorkStation();
    return 0;
}
```

#### Важные нюансы:

* Эффект: Программный вызов дает ровно тот же результат, что и нажатие клавиш Win + L.
* Асинхронность: Функция возвращает управление сразу после того, как запрос на блокировку был инициирован. Она не ждет, пока экран фактически станет черным.
* Ограничения: Функция сработает только если процесс запущен в контексте интерактивного рабочего стола (т.е. обычного пользователя). Она может не сработать из-под системной службы (Windows Service) без дополнительных манипуляций с сессиями.
* Обратное действие: Программно разблокировать компьютер (ввести пароль за пользователя) стандартными средствами API Windows нельзя в целях безопасности. [1, 2, 10, 11, 12] 

#### Ссылки

- [1] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-lockworkstation)
- [2] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-lockworkstation)
- [3] [https://www.itarian.com](https://www.itarian.com/blog/how-to-lock-pc-screen/)
- [4] [https://www.itarian.com](https://www.itarian.com/blog/how-to-lock-pc-screen/)
- [5] [https://www.itprotoday.com](https://www.itprotoday.com/microsoft-windows/how-can-i-lock-a-workstation-from-the-command-line-)
- [6] [https://stackoverflow.com](https://stackoverflow.com/questions/20733441/lock-windows-workstation-using-python)
- [7] [https://medium.com](https://medium.com/codex/auto-locking-pc-when-youre-away-with-python-e184fe2b4bd5)
- [8] [https://stackoverflow.com](https://stackoverflow.com/questions/38595976/programmatically-lock-workstation-from-windows-service)
- [9] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/windows/win32/shutdown/how-to-lock-the-workstation)
- [10] [https://www.reddit.com](https://www.reddit.com/r/computer/comments/15c2llw/is_there_a_way_to_lock_my_computer_station_so/?tl=ru#:~:text=%D0%9D%D0%B0%D0%B6%D0%BC%D0%B8%D1%82%D0%B5%20%D0%BA%D0%BB%D0%B0%D0%B2%D0%B8%D1%88%D1%83%20Windows%20+L%20%D0%B8%D0%BB%D0%B8%20CTRL+ALT+DEL.%20%D0%AD%D1%82%D0%BE,%D0%BE%D1%82%20%D0%BF%D1%80%D0%BE%D1%81%D0%BC%D0%BE%D1%82%D1%80%D0%B0%20%D1%8D%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B8%D0%BB%D0%B8%20%D0%B2%D1%85%D0%BE%D0%B4%D0%B0%20%D0%B2%20%D1%81%D0%B8%D1%81%D1%82%D0%B5%D0%BC%D1%83.)
- [11] [https://www.jasinskionline.com](http://www.jasinskionline.com/windowsapi/ref/l/lockworkstation.html)
- [12] [https://learn.microsoft.com](https://learn.microsoft.com/ru-ru/windows/win32/api/winuser/nf-winuser-lockworkstation#:~:text=%D0%9A%D0%BE%D0%BC%D0%BC%D0%B5%D0%BD%D1%82%D0%B0%D1%80%D0%B8%D0%B8%20%D0%A4%D1%83%D0%BD%D0%BA%D1%86%D0%B8%D1%8F%20LockWorkStation%20%D0%B2%D1%8B%D0%B7%D1%8B%D0%B2%D0%B0%D0%B5%D1%82%D1%81%D1%8F%20%D1%82%D0%BE%D0%BB%D1%8C%D0%BA%D0%BE%20%D0%BF%D1%80%D0%BE%D1%86%D0%B5%D1%81%D1%81%D0%B0%D0%BC%D0%B8%2C%20%D0%B7%D0%B0%D0%BF%D1%83%D1%89%D0%B5%D0%BD%D0%BD%D1%8B%D0%BC%D0%B8,%D1%81%D1%82%D0%BE%D0%BB%D0%B5%20%D0%B8%D0%BB%D0%B8%20%D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%20%D0%BE%D1%82%D0%BA%D0%BB%D0%BE%D0%BD%D1%8F%D0%B5%D1%82%D1%81%D1%8F%20%D0%B1%D0%B8%D0%B1%D0%BB%D0%B8%D0%BE%D1%82%D0%B5%D0%BA%D0%BE%D0%B9%20DLL%20GINA.)


### Отслеживание гашения экрана в Windows

Да, это можно сделать. Однако важно понимать: Windows разделяет понятия «блокировка системы» и «выключение экрана». Экран может погаснуть из-за простоя, даже если компьютер не заблокирован (например, если вы смотрите видео или просто отошли, а автоблокировка выключена).
Вот основные способы программно поймать этот момент:

#### 1. Подписка на уведомления о питании (WinAPI)
Это самый современный и точный способ. Вам нужно зарегистрировать окно для получения уведомлений об изменении состояния дисплея через функцию RegisterPowerSettingNotification.

* Используемый GUID: GUID_CONSOLE_DISPLAY_STATE (или GUID_MONITOR_POWER_ON для более старых систем).
* Сообщение: При изменении состояния экрана ваше окно получит сообщение WM_POWERBROADCAST с параметром PBT_POWERSETTINGCHANGE.
* Значения в данных:
* 0 — экран выключен.
   * 1 — экран включен.
   * 2 — экран работает в режиме пониженной яркости (dimmed). [1, 2, 3, 4] 

#### 2. Отслеживание времени простоя (GetLastInputInfo)
Если вам не нужно событие от самой системы, вы можете сами вычислить, когда экран должен погаснуть:

   1. Используйте функцию GetLastInputInfo, чтобы узнать время последнего действия пользователя (движение мыши, нажатие клавиш).
   2. Сравните это время с текущим системным временем.
   3. Запросите через API настройки электропитания пользователя (таймаут отключения экрана) и, если время простоя превысило этот порог, считайте, что экран погас. [5, 6, 7] 

#### 3. Использование событий системы (Event Log)
Windows записывает изменения состояния в журнал событий. Это полезно, если нужно проверить факт отключения постфактум:

* Журнал: System.
* Источник: Power-Troubleshooter или Kernel-Power.
* Событие 4802: Генерируется, когда запускается экранная заставка (часто совпадает с моментом гашения экрана при соответствующих настройках). [8, 9, 10] 

#### 4. Для .NET (C#)
В среде .NET проще всего использовать событие SystemEvents.PowerModeChanged, однако оно чаще реагирует на переход в спящий режим, а не на простое выключение монитора. Для точного отслеживания именно монитора лучше использовать P/Invoke для вызова RegisterPowerSettingNotification, как описано в первом пункте. [1] 
Важный нюанс: На современных ноутбуках с режимом Modern Standby экран гаснет практически одновременно с переходом системы в режим низкого энергопотребления, что может потребовать использования RegisterSuspendResumeNotification. [11, 12] 

#### Ссылки

- [1] [https://stackoverflow.com](https://stackoverflow.com/questions/79889216/how-to-get-the-values-of-wm-powerbroadcast-pbt-powersettingchange-when-monitor#:~:text=Protected%20Overrides%20Sub%20WndProc%28ByRef%20m%20As%20Message%29,End%20If%20End%20If%20MyBase.WndProc%28m%29%20End%20Sub.)
- [2] [https://forum.lazarus.freepascal.org](https://forum.lazarus.freepascal.org/index.php?topic=57125.0)
- [3] [https://github.com](https://github.com/MicrosoftDocs/win32/blob/docs/desktop-src/Power/pbt-powersettingchange.md)
- [4] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/windows/win32/power/registering-for-power-events)
- [5] [https://stackoverflow.com](https://stackoverflow.com/questions/8820615/how-to-check-in-c-if-the-system-is-active)
- [6] [https://blog.aaronlenoir.com](https://blog.aaronlenoir.com/2016/02/16/detect-system-idle-in-windows-applications-2/)
- [7] [https://dzone.com](https://dzone.com/articles/detect-if-user-isnbspidle)
- [8] [https://superuser.com](https://superuser.com/questions/1258473/display-all-power-related-events-turn-on-off-sleep-hibernate#:~:text=Open%20the%20Windows%20System%20Log%2C%20choose%20Filter,option%22.%20However%2C%20you%20can%20make%20it%20faster:)
- [9] [https://www.manageengine.com](https://www.manageengine.com/products/active-directory-audit/kb/logon-logoff-events/event-id-4802.html)
- [10] [https://community.spiceworks.com](https://community.spiceworks.com/t/windows-event-ids-4802-and-4803-missing-in-windows-11/929428)
- [11] [https://community.osr.com](https://community.osr.com/t/modern-standby-and-power-off-the-display-s/57759)
- [12] [https://learn.microsoft.com](https://learn.microsoft.com/en-us/answers/questions/48395/cannot-receive-wm-powerbroadcast-message-on-win10)

