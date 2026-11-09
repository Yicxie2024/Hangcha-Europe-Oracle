"""Per-user, Windows Task Scheduler PREVIEW-ONLY. Never schedules Excel writes."""
from __future__ import annotations
import os
from datetime import datetime
from pathlib import Path

TASK_NAME = 'Hangcha_Weekly_Excel_Autofill'  # replace unsafe legacy task on approval
BASE = Path(__file__).resolve().parent


def _folder():
    if os.name != 'nt':
        raise RuntimeError('Windows 任务计划程序仅支持 Windows')
    import win32com.client
    service = win32com.client.Dispatch('Schedule.Service')
    service.Connect()
    return service, service.GetFolder('\\')


def enable_preview_only():
    """Caller MUST show user a positive confirmation before invoking."""
    service, folder = _folder()
    task = service.NewTask(0)
    task.RegistrationInfo.Description = '每日工作日16:45打开 GUI，执行只读扫描和异常预览；绝不自动写入 Excel。'
    task.Settings.Enabled = True
    task.Settings.StartWhenAvailable = True
    task.Settings.DisallowStartIfOnBatteries = False
    task.Settings.StopIfGoingOnBatteries = False
    task.Settings.ExecutionTimeLimit = 'PT30M'
    task.Settings.MultipleInstances = 2
    task.Principal.LogonType = 3
    trigger = task.Triggers.Create(3)
    trigger.DaysOfWeek = 2 + 4 + 8 + 16 + 32
    trigger.WeeksInterval = 1
    trigger.StartBoundary = datetime.now().strftime('%Y-%m-%dT16:45:00')
    trigger.Enabled = True
    action = task.Actions.Create(0)
    pythonw = BASE / '.venv' / 'Scripts' / 'pythonw.exe'
    pythonexe = BASE / '.venv' / 'Scripts' / 'python.exe'
    action.Path = str(pythonw if pythonw.is_file() else pythonexe)
    action.Arguments = f'"{BASE / "gui.py"}" --scheduled-preview'
    action.WorkingDirectory = str(BASE)
    folder.RegisterTaskDefinition(TASK_NAME, task, 6, '', '', 3)


def disable():
    _, folder = _folder()
    folder.DeleteTask(TASK_NAME, 0)


def status():
    try:
        _, folder = _folder()
        task = folder.GetTask(TASK_NAME)
        # Older versions used this task name for UNATTENDED APPLY. Expose it.
        actions = [task.Definition.Actions.Item(i).Arguments for i in range(1, task.Definition.Actions.Count+1)]
        unsafe = any('--apply' in str(args) for args in actions)
        return 'legacy_write' if unsafe else ('enabled' if task.Enabled else 'disabled')
    except Exception:
        return 'none'
