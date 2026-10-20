<a id="top"></a>

<p align="center">
  <img src="hik.png" width="300" alt="HIK-Report">
</p>

<h1 align="center">HIK-Report</h1>

<p align="center">
  <b>Login &amp; activity reports for Hikvision NVRs — straight from the terminal into a polished Excel file.</b><br>
  <sub>Отчёты по входам и активности пользователей Hikvision — прямо из терминала в красивый Excel.</sub>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%20%2F%2011-0078D6?style=flat-square&logo=windows&logoColor=white">
  <img src="https://img.shields.io/badge/python-3.x-3776AB?style=flat-square&logo=python&logoColor=white">
  <img src="https://img.shields.io/badge/single-.exe%20(no%20install)-2ea44f?style=flat-square">
  <img src="https://img.shields.io/badge/API-Hikvision%20ISAPI-E60012?style=flat-square">
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square">
</p>

<p align="center">
  <a href="#english"><img src="https://img.shields.io/badge/🇬🇧%20Read%20in-English-2ea44f?style=for-the-badge"></a>
  &nbsp;
  <a href="#russian"><img src="https://img.shields.io/badge/🇷🇺%20Читать%20на-Русском-1f6feb?style=for-the-badge"></a>
</p>

---

<a id="english"></a>

## 🇬🇧 English

**HIK-Report** connects to Hikvision NVRs over **ISAPI** and builds a clean Excel report of a
selected user's **logins/logouts** (times, source IP, session duration) and their **actions**
(PTZ control and archive playback — with camera names), grouped under each session.

Everything happens **in the terminal** — no browser, no web interface, just HTTP requests.
It ships as a **single `.exe`** that runs on a bare Windows 10/11 machine with **nothing to install**.

### ✨ Features

- 📋 **Logins/logouts** per user — date, time, source IP and session duration.
- 🎥 **Actions under each session** — PTZ and archive playback, shown as `Dxx — Camera name`.
- 🧭 **Camera names** resolved automatically from the NVR.
- 📊 **Polished Excel** output (styled header, colour-coded rows, summary).
- 🌈 Smooth **rainbow progress bar** and blinking **SUCCESS / ERROR** badges.
- ⚙️ **Settings menu** (press `0`) — add / edit / delete recorders: internal IP &amp; port,
  external host &amp; port.
- 💾 **Config stored inside the `.exe` itself** (NTFS alternate data stream) — no side files.
- 📦 **Zero dependencies for the user** — Python and all libraries are bundled in the `.exe`.

### 🚀 Quick start

1. Download **`HIK-Report.exe`** and run it (double-click).
2. On first launch Windows SmartScreen may warn about an unknown publisher →
   **More info → Run anyway** (the file is unsigned; that's normal for a self-built tool).
3. The list of recorders is **empty on first run** — press **`0`** to open **Settings** and add one
   (internal IP + port, optionally an external host + port for remote access).
4. Pick a recorder, enter **login** (`Tab` jumps to password) and **password**.
5. Choose a **user** and a **date range** (`dd.mm.yyyy`, empty = today).
6. The Excel report is saved **next to the program** and opens automatically.
   At the end: **ESC** to quit, **Enter** for a new report.

### ⚙️ Settings &amp; persistence

Press **`0`** in the recorder menu to manage recorders. Settings are saved **inside the `.exe`**
using an NTFS **alternate data stream**, so there are no extra config files and the settings
travel with the executable when it is copied **on an NTFS drive**.

> ⚠️ The stored settings are lost when the `.exe` is copied to an **exFAT/FAT** drive (most USB
> sticks), put into a **ZIP**, or **emailed** — in those cases it starts again with an empty list.
> (Keeping data literally inside a single running `.exe` across such transfers is not technically
> possible — this is the one unavoidable trade-off.)

### 📄 What's in the report

| Column | Meaning |
|---|---|
| Login / Logout | Session start and end (date &amp; time) |
| Duration | `logout − login`, matched per source IP |
| IP address | Where the user connected from |
| ↳ Actions | PTZ / playback during that session, with `Dxx — Camera name` |

> `—` in *Logout/Duration* means the NVR never recorded a logout (session closed by
> timeout/disconnect), so the exact duration cannot be known.
> **Live view is not logged by Hikvision NVRs** and therefore cannot be reported.

### 🛠️ Build from source

```bash
pip install requests openpyxl colorama pyinstaller
pyinstaller --onefile --console --name "HIK-Report" --icon hik.ico nvr_report.py
```

The result is `dist/HIK-Report.exe`. You can also run the script directly with Python
(`python nvr_report.py`) if you have the dependencies installed.

### 🔒 Security &amp; privacy

- **No addresses are hard-coded** — you add your own recorders at runtime.
- Credentials are typed at runtime and used only for the live HTTP request; **they are not stored**.
- Use this tool **only on systems you own or are authorised to audit.**

### 📜 License

Released under the **MIT License** — see [`LICENSE`](LICENSE).
Free to use, modify and distribute; just keep the copyright notice.

<p align="right"><a href="#top">↑ back to top</a></p>

---

<a id="russian"></a>

## 🇷🇺 Русский

**HIK-Report** подключается к регистраторам **Hikvision** по **ISAPI** и формирует аккуратный
Excel-отчёт по выбранному пользователю: **входы/выходы** (время, IP, длительность сессии) и его
**действия** (управление PTZ и воспроизведение архива — с названиями камер), сгруппированные
под каждой сессией.

Всё происходит **в терминале** — без браузера и веб-интерфейса, только HTTP-запросы.
Поставляется как **один `.exe`**, который работает на «голой» Windows 10/11 — **ставить ничего не нужно**.

### ✨ Возможности

- 📋 **Входы/выходы** по пользователю — дата, время, IP источника и длительность сессии.
- 🎥 **Действия под каждой сессией** — PTZ и воспроизведение архива в виде `Dxx — Название камеры`.
- 🧭 **Названия камер** подтягиваются с регистратора автоматически.
- 📊 **Оформленный Excel** (стильная шапка, цветные строки, сводка).
- 🌈 Плавный **радужный прогресс-бар** и мигающие бейджи **SUCCESS / ERROR**.
- ⚙️ **Меню настроек** (клавиша `0`) — добавить / изменить / удалить регистраторы:
  внутренний IP и порт, внешний хост и порт.
- 💾 **Настройки хранятся внутри самого `.exe`** (NTFS ADS) — без отдельных файлов.
- 📦 **Никаких зависимостей у пользователя** — Python и все библиотеки уже внутри `.exe`.

### 🚀 Быстрый старт

1. Скачайте **`HIK-Report.exe`** и запустите (двойной клик).
2. При первом запуске Windows SmartScreen может предупредить →
   **Подробнее → Выполнить в любом случае** (файл без подписи — это нормально для самодельной утилиты).
3. Список регистраторов **при первом запуске пуст** — нажмите **`0`**, откройте **Настройки** и добавьте
   регистратор (внутренний IP + порт, при необходимости внешний хост + порт для удалёнки).
4. Выберите регистратор, введите **логин** (`Tab` — переход к паролю) и **пароль**.
5. Выберите **пользователя** и **диапазон дат** (`дд.мм.гггг`, пусто = сегодня).
6. Excel-отчёт сохранится **рядом с программой** и откроется сам.
   В конце: **ESC** — выход, **Enter** — новый отчёт.

### ⚙️ Настройки и хранение

Клавиша **`0`** в меню регистраторов открывает управление списком. Настройки сохраняются
**внутри самого `.exe`** в альтернативном потоке NTFS (ADS) — отдельных файлов нет, а настройки
«переезжают» вместе с exe при копировании **по диску NTFS**.

> ⚠️ Настройки теряются при копировании exe на диск **exFAT/FAT** (большинство флешек),
> упаковке в **ZIP** или отправке **по почте** — тогда список снова будет пустым.
> (Хранить данные буквально внутри одного запущенного `.exe` при таких переносах технически
> невозможно — это единственный неизбежный компромисс.)

### 📄 Что в отчёте

| Колонка | Значение |
|---|---|
| Вход / Выход | Начало и конец сессии (дата и время) |
| Длительность | `выход − вход`, сопоставлено по одному IP |
| IP-адрес | Откуда подключался пользователь |
| ↳ Действия | PTZ / воспроизведение в этой сессии, в виде `Dxx — Название камеры` |

> `—` в *Выход/Длительность* означает, что регистратор не записал выход (сессия закрылась по
> таймауту/разрыву) — точную длительность определить нельзя.
> **Живой просмотр камер регистраторы Hikvision не логируют**, поэтому его в отчёте нет.

### 🛠️ Сборка из исходников

```bash
pip install requests openpyxl colorama pyinstaller
pyinstaller --onefile --console --name "HIK-Report" --icon hik.ico nvr_report.py
```

Результат — `dist/HIK-Report.exe`. Можно запускать и сам скрипт (`python nvr_report.py`),
если зависимости установлены.

### 🔒 Безопасность и приватность

- **Никаких адресов в коде** — регистраторы вы добавляете сами при запуске.
- Логин и пароль вводятся в момент работы и используются только для живого HTTP-запроса;
  **они нигде не сохраняются**.
- Используйте утилиту **только на системах, которыми владеете или имеете право проверять.**

### 📜 Лицензия

Распространяется под лицензией **MIT** — см. [`LICENSE`](LICENSE).
Можно свободно использовать, изменять и распространять; нужно лишь сохранить строку об авторстве.

<p align="right"><a href="#top">↑ наверх</a></p>
