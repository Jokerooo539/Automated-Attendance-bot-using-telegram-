# Attendance bot: setup on the old Motorola (Termux)

## 1. Telegram (on your main phone)
1. Open @BotFather, send `/newbot`, copy the **token**.
2. Send `/setprivacy`, choose your bot, then **Disable** (so it can read group messages).
3. Create or use 3 Telegram groups (Block A, B, C). Add the bot to each one.
   If a group already had the bot, remove it and add it again after changing privacy.
4. Block leaders join their group.

## 2. Install on the old phone
1. Install **Termux**, **Termux:Boot** and (optional) **Termux:API** from **F-Droid** (not the Play Store).
2. In Termux:
   ```
   pkg update && pkg upgrade
   pkg install python
   pip install requests openpyxl
   mkdir -p ~/attendance && cd ~/attendance
   ```
3. Run `termux-setup-storage` once and tap **Allow**. Copy `attendance_bot.py` and `config.json` into `~/attendance`
   (from `~/storage/downloads/`). Put your `daily_attendance.xlsx` in the phone's storage folder
   `/storage/emulated/0/Attendance/` (see section 9).
4. Copy `config.example.json` to `config.json` and edit it (token, admin id, group ids).

## 3. Find the ids
Run `python attendance_bot.py` once. In each group send `/id`; the bot replies with the group id
(a negative number like -100123...). In a private chat with the bot send `/id` to get **your** user id
for `admins`. Put them in `config.json`, then stop (Ctrl+C) and start again.

## 4. Test first
Put a **copy** of the Excel file at the `xlsx_path` from `config.json` and post a sample message in a test group.
Check the file with `/file` before using your real sheet.

## 5. Keep it running 24/7
- Phone: Settings > Apps > Termux and Termux:Boot > Battery > **Unrestricted**. Keep it on the charger.
- Lock Termux in the recent-apps screen so Android doesn't close it.
- Create the auto-start script:
  ```
  mkdir -p ~/.termux/boot
  cat > ~/.termux/boot/start-bot.sh <<'X'
  #!/data/data/com.termux/files/usr/bin/sh
  termux-wake-lock
  cd ~/attendance
  nohup python attendance_bot.py >> bot.log 2>&1 &
  X
  chmod +x ~/.termux/boot/start-bot.sh
  ```
- Open Termux:Boot once so Android allows it. After a reboot the bot starts by itself.
- Check it any time by sending `/status` to the bot. Log file: `~/attendance/bot.log`.

## 6. Daily use
- Leaders post as usual. The bot saves it and replies with a summary.
- A wrong post: the leader edits the message (the bot updates the cell) or reposts it.
- Admin commands: `/file` (get the Excel), `/summary` or `/summary 02/10/2026` (the combined message for NI Attendance), `/status`, and `/add Rohit Kumar, B, 05/10/2026` (adds a new worker to that block; the date is optional).
- In Termux you can also run: `python attendance_bot.py add "Rohit Kumar" B 05/10/2026`.
- At `daily_file_time` the bot sends the Excel file to the admins as an off-phone backup.
- **The Excel is in the phone's normal storage** (Files app > Attendance), so you can open and edit it directly. See section 9.
- When someone leaves, write **LEFT** in their cell for the first day after their last day.
  The bot then stops listing them as "Not reported", and the next month's tab drops them.

## 7. Nightly combined post (v2)
1. Create a Telegram group for the combined message (e.g. "NI Attendance") and add the bot to it.
2. Send `/id` in that group and copy the chat id (it starts with a minus sign).
3. Put it in `config.json` as `"main_chat": -100...` (a number, no quotes) and set `"combined_post_time": "23:30"`.
4. Restart the bot (`pkill -f attendance_bot.py`; the boot loop starts it again in 10 seconds).
5. Every night at that time the bot posts the combined message for **today's date** to the main group,
   in the same format you use now (date, then 🔴 Block A 🔴 with each name and code).
   If nothing was saved that day, it posts nothing and sends the admins a short note.
6. `/post` (admin) sends today's combined message to the main group right now; `/post 02/10/2026` for another date.
7. It posts only once a day, even if the phone restarts. The bot remembers this in `state.json`.

   Even without a main group, the **admins** get the same combined message every night, plus a warning line
   for any block that posted nothing that day (`"send_combined_to_admins": false` turns this off).
   If the phone was off at the post time, the bot posts yesterday's message as soon as it is back and has caught up.

## 8. When the internet or the phone goes down (v2.1)
- **Internet lost:** the bot keeps running, retries every 10 to 60 seconds, and logs "network problem".
- **Back online:** Telegram has kept every message that was sent meanwhile. The bot reads them in order, saves them to
  Excel, replies in the groups, and tells the admins: "Back online after N min. Caught up on K message(s)".
- **Phone off or Termux killed:** same thing after it starts again. Telegram keeps messages for **24 hours only**.
  If the bot was off longer, it warns the admins when it starts. Use `/load` for anything older.
- **/load (admin):** type `/load B`, press Enter for a new line, paste the block's attendance message, send.
  The bot saves it exactly as if the leader had posted it.
- **One copy only:** the bot refuses to start twice (lock file), so a second start can't cause double replies.
- **Watchdog** (restarts the bot if Android kills Termux during the day, not just at boot). Needs the **Termux:API** app
  (F-Droid) and `pkg install termux-api`:
  ```
  printf '#!/data/data/com.termux/files/usr/bin/sh\nif ! pgrep -f "start-bot.sh|attendance_bot.py" >/dev/null; then\n  nohup sh /data/data/com.termux/files/home/.termux/boot/start-bot.sh >/dev/null 2>&1 &\nfi\n' > ~/check-bot.sh && chmod +x ~/check-bot.sh && termux-job-scheduler --script ~/check-bot.sh --period-ms 900000 --persisted true
  ```
  It checks every 15 minutes. Test it: `pkill -f start-bot.sh; pkill -f attendance_bot.py`, wait up to 15 minutes, then send `/status`.
- **Start now without rebooting:** `nohup sh ~/.termux/boot/start-bot.sh >/dev/null 2>&1 &`

## 9. Where the Excel file lives (v2.5)
The Excel file is in the phone's normal storage, so you can open it any time:

- **File:** `/storage/emulated/0/Attendance/daily_attendance.xlsx` (Files app > Internal storage > **Attendance**)
- **Daily backups:** `Attendance/backups/`, one copy per day, made before the first change of that day
- **The bot's own files** (lock, state, offset, `bot.log`) stay in `~/attendance`, so they never clutter your folder

**Opening and editing it**
- Open it from the Files app with WPS Office, Excel or any spreadsheet app, change what you need, **save, and close the app**.
- The bot reads the file again for every message, so your edits are kept and the bot's entries are kept.
- Don't leave the file open in the app for hours. If the app saves later using the old copy it opened, it can overwrite
  what the bot saved in between. Open, change, save, close.
- Don't rename the file, move it, or rename the month tabs (AUG, SEP, OCT) or the "Block A/B/C" and "Name" rows. The bot finds workers by them.
- `/file` still sends you the current file in Telegram, and `/status` shows its path and when it was last saved.

**If the file is missing or half-saved** (for example another app was cut off while saving)
- The bot tells the admins, keeps every message waiting, and saves them automatically once the file works again
  (Telegram holds messages for 24 hours). It tells you when it is back to normal.
- To fix it: open the file and save it again, or copy the newest file from `Attendance/backups/` over it.

**One-time move from the old location** (run in Termux)
1. Put the new `attendance_bot.py` in your Download > attendance_bot folder, then:
   ```
   mkdir -p /storage/emulated/0/Attendance
   cp ~/storage/downloads/attendance_bot/attendance_bot.py ~/attendance/
   cp ~/attendance/daily_attendance.xlsx /storage/emulated/0/Attendance/
   mv ~/attendance/daily_attendance.xlsx ~/attendance/old_daily_attendance.xlsx
   cd ~/attendance
   python -c "import json;c=json.load(open('config.json'));c['xlsx_path']='/storage/emulated/0/Attendance/daily_attendance.xlsx';c['backup_dir']='/storage/emulated/0/Attendance/backups';c['data_dir']='/data/data/com.termux/files/home/attendance';c.pop('drive',None);json.dump(c,open('config.json','w'),indent=2);print('saved')"
   pkill -f attendance_bot.py
   ```
   The boot script starts the bot again within 10 seconds.
2. Send `/status`. It should show `Bot v2.5 running` and the new Excel path.
3. The old copy is kept as `~/attendance/old_daily_attendance.xlsx` and is no longer used. Delete it once you are happy.

## 10. Fixing a forgotten amount (v2.3)
- **Quickest:** `/advance Bablu 300` (today), `/advance Bablu 300 02/10/2026` (another date). It keeps the worker's code and
  adds the amount to what is already there: `P1` + `500` becomes `P1` / `500+300`.
- If the worker has nothing saved that day, add the code: `/advance Bablu 300 02/10/2026 P1`.
- If two workers share a first name it asks for the full name. Write `2000+500` to add two amounts at once.
- To correct a wrong amount, repost the worker's line (`Bablu p1 500`) in the group or with `/load B`.

## 11. Safety notes
- Admin commands (`/file`, `/summary`, `/post`, `/add`, `/advance`, `/load`, `/status`) work **only in a private chat** with the bot,
  so the Excel file can never be sent into a group by mistake.
- The bot token is masked in `bot.log`, so you can paste the log when asking for help.
- A message that contains two different dates is rejected with a warning instead of being saved under the wrong day.
- Ticks, emojis, list numbers ("1. Rahul") and "Rs" / "₹" / "/-" after a code are ignored when reading messages.

