from __future__ import annotations

import json
import os
import re
import shlex
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

from .cy_adapter import CYAdapter


class CYMacOSAdapter(CYAdapter):
    def __init__(self, app_name: str | None = None):
        self.app_name = app_name or os.environ.get('CY_APP_NAME', 'CYCardScanner')
        self.bundle_name = os.environ.get('CY_APP_NAME', self.app_name)
        self.layout = self._load_layout()
        self.cliclick_path = self._resolve_tool('cliclick')
        self.tesseract_path = self._resolve_tool('tesseract')
        self.scroll_tool = Path(__file__).resolve().parents[2] / 'scripts' / 'macos' / 'cgscroll'
        self.needs_dialog_cleanup = False

    def _resolve_tool(self, name: str) -> str:
        for candidate in (
            shutil.which(name),
            f'/opt/homebrew/bin/{name}',
            f'/usr/local/bin/{name}',
            f'/usr/bin/{name}',
        ):
            if candidate and Path(candidate).exists():
                return candidate
        raise RuntimeError(f'Missing required command: {name}. Install it or add it to PATH.')

    def _load_layout(self) -> dict:
        env_path = os.environ.get('CY_LAYOUT_PATH', '').strip()
        if env_path:
            candidate = Path(env_path)
        else:
            candidate = Path(__file__).resolve().parent / 'cy_layout.json'

        if candidate.exists():
            return json.loads(candidate.read_text(encoding='utf-8'))

        example = Path(__file__).resolve().parent / 'cy_layout.example.json'
        if example.exists():
            return json.loads(example.read_text(encoding='utf-8'))

        raise RuntimeError('Missing CY layout config. Create backend/cy_layout.json from cy_layout.example.json')

    def _timing(self, key: str, default_ms: int) -> float:
        return (self.layout.get('timing_ms', {}).get(key, default_ms)) / 1000.0

    def _point(self, key: str) -> tuple[int, int]:
        value = self.layout.get('points', {}).get(key)
        if not value or len(value) != 2:
            raise RuntimeError(f'Missing CY layout point: {key}')
        return int(value[0]), int(value[1])

    def _region(self, key: str) -> tuple[int, int]:
        value = self.layout.get('regions', {}).get(key)
        if not value or len(value) != 2:
            raise RuntimeError(f'Missing CY layout region point: {key}')
        return int(value[0]), int(value[1])

    def _osascript(self, script: str) -> str:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=20,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or 'osascript failed')
        return (result.stdout or '').strip()

    def _is_app_running(self) -> bool:
        script = f'tell application "System Events" to (name of processes) contains "{self.bundle_name}"'
        return self._osascript(script).lower() == 'true'

    def _activate_app(self):
        self._osascript(f'tell application "{self.app_name}" to activate')
        time.sleep(self._timing('after_activate', 250))

    def _position_window_right(self):
        script = f'''
        tell application "Finder"
            set screenBounds to bounds of window of desktop
            set screenWidth to item 3 of screenBounds
            set screenHeight to item 4 of screenBounds
        end tell

        set targetTop to {int(self.layout.get('window', {}).get('top', 33))}
        set targetHeight to {int(self.layout.get('window', {}).get('height', 991))}

        tell application "{self.app_name}"
            activate
        end tell

        tell application "System Events"
            tell process "{self.bundle_name}"
                set frontmost to true
                repeat 20 times
                    if exists window 1 then exit repeat
                    delay 0.1
                end repeat
                if exists window 1 then
                    set position of window 1 to {{screenWidth / 2, targetTop}}
                    set size of window 1 to {{screenWidth / 2, targetHeight}}
                    delay 0.1
                    set windowPosition to position of window 1
                    set windowSize to size of window 1
                    return ((item 1 of windowPosition) as string) & "," & ((item 2 of windowPosition) as string) & "," & ((item 1 of windowSize) as string) & "," & ((item 2 of windowSize) as string)
                end if
            end tell
        end tell
        return ""
        '''
        bounds = self._osascript(script)
        time.sleep(self._timing('after_activate', 250))
        return bounds

    def _ensure_app_running(self):
        was_running = self._is_app_running()
        self._activate_app()
        bounds = self._position_window_right()
        return was_running, bounds

    def _cliclick(self, command: str):
        result = subprocess.run([self.cliclick_path, command], capture_output=True, text=True, timeout=10)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or f'cliclick failed: {command}')

    def _click_point(self, key: str):
        x, y = self._point(key)
        self._cliclick(f'c:{x},{y}')
        time.sleep(self._timing('after_click', 150))

    def _wheel_scroll(self, delta: int | str, steps: int):
        if steps <= 0:
            return {'delta': '0', 'steps': 0}
        if not self.scroll_tool.exists():
            raise RuntimeError(f'Missing scroll helper: {self.scroll_tool}')
        delta = str(delta)
        result = subprocess.run(
            [str(self.scroll_tool), delta, str(steps)],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or 'CoreGraphics wheel scroll failed')
        time.sleep(0.15)
        return {'delta': delta, 'steps': steps}

    def _wheel_scroll_down(self, steps: int):
        delta = self.layout.get('scroll', {}).get('wheel_delta', -12)
        return self._wheel_scroll(delta, steps)

    def _reset_slab_table_to_top(self):
        scroll = self.layout.get('scroll', {})
        return self._wheel_scroll(
            scroll.get('reset_delta', 12),
            int(scroll.get('reset_steps', 12)),
        )

    def _paste_text(self, text: str):
        script = f'''
        tell application "System Events"
            keystroke "a" using command down
            key code 51
            keystroke {json.dumps(str(text))}
        end tell
        '''
        self._osascript(script)
        time.sleep(self._timing('after_paste', 120))

    def _normalize_slab_type(self, slab_type: str) -> str:
        value = str(slab_type or '').strip().upper()
        aliases = {
            'PSA': 'PSA',
            'CGC': 'CGC',
            'BGS': 'BGS',
            'SGC': 'SGC',
        }
        if value not in aliases:
            raise ValueError(f'Unsupported slab type: {slab_type}')
        return aliases[value]

    def _ensure_manual_entry(self):
        # Try the measured tab location twice, with a small settle delay,
        # because the app may still be animating after activation/resize.
        self._click_point('manual_entry_tab')
        time.sleep(0.2)
        self._click_point('manual_entry_tab')
        time.sleep(0.2)

    def _select_slab_type(self, slab_type: str):
        slab = self._normalize_slab_type(slab_type)
        scroll_steps = {
            'PSA': 0,
            'CGC': 1,
            'BGS': 2,
            'SGC': 3,
        }
        self._click_point('slab_list_center')
        reset_debug = self._reset_slab_table_to_top()
        scroll_debug = self._wheel_scroll_down(
            int(self.layout.get('scroll', {}).get(f'{slab.lower()}_reveal_clicks', scroll_steps.get(slab, 0)))
        )
        self._click_point(f'slab_{slab.lower()}')
        return {
            'reset': reset_debug,
            'target': scroll_debug,
            'clicked': f'slab_{slab.lower()}',
        }

    def _dismiss_possible_error_dialog(self):
        dismissed = []
        dialog_pattern = re.compile(
            r'error|no card found|purchase price|certificate number|failed to fetch|card details|cannot fetch|could not fetch',
            re.IGNORECASE,
        )
        dialog_text = ''
        dialog_detected = False
        final_dialog_text = ''
        ok_x, ok_y = self._point('error_ok_button')
        ok_points = [
            (ok_x, ok_y),
            (ok_x - 40, ok_y),
            (ok_x + 40, ok_y),
            (ok_x, ok_y - 10),
            (ok_x, ok_y + 10),
        ]

        for attempt in range(3):
            try:
                dialog_text, _, _ = self._read_region_ocr('error_dialog_top_left', 'error_dialog_bottom_right', 'cy-dialog')
            except Exception:
                dialog_text = ''

            final_dialog_text = dialog_text
            if not dialog_pattern.search(dialog_text or ''):
                return {
                    'dismissed': dismissed,
                    'dialog_detected': dialog_detected,
                    'dialog_text': final_dialog_text,
                }

            dialog_detected = True
            for point_index, (x, y) in enumerate(ok_points):
                try:
                    self._cliclick(f'c:{x},{y}')
                    dismissed.append(f'error_ok_button_{attempt + 1}_{point_index + 1}')
                    time.sleep(0.12)
                except Exception:
                    pass
            for key in ('return', 'esc'):
                try:
                    self._cliclick(f'kp:{key}')
                    dismissed.append(f'{key}_{attempt + 1}')
                    time.sleep(0.12)
                except Exception:
                    pass
            time.sleep(0.25)

        return {
            'dismissed': dismissed,
            'dialog_detected': dialog_detected,
            'dialog_text': final_dialog_text,
        }

    def _read_front_window_text(self) -> str:
        script = f'''
        tell application "System Events"
            tell process "{self.bundle_name}"
                set outText to ""
                try
                    set staticVals to value of static texts of front window
                    repeat with v in staticVals
                        set outText to outText & (v as string) & linefeed
                    end repeat
                end try
                try
                    set fieldVals to value of text fields of front window
                    repeat with v in fieldVals
                        set outText to outText & (v as string) & linefeed
                    end repeat
                end try
                try
                    set groupCount to count of groups of front window
                    repeat with i from 1 to groupCount
                        try
                            set nestedStatics to value of static texts of group i of front window
                            repeat with v in nestedStatics
                                set outText to outText & (v as string) & linefeed
                            end repeat
                        end try
                        try
                            set nestedFields to value of text fields of group i of front window
                            repeat with v in nestedFields
                                set outText to outText & (v as string) & linefeed
                            end repeat
                        end try
                    end repeat
                end try
                return outText
            end tell
        end tell
        '''
        return self._osascript(script)

    def _read_region_ocr(self, top_left_key: str, bottom_right_key: str, prefix: str) -> tuple[str, str, str]:
        left, top = self._region(top_left_key)
        right, bottom = self._region(bottom_right_key)
        width = max(1, right - left)
        height = max(1, bottom - top)
        debug_dir = Path.home() / 'live-comps-debug'
        debug_dir.mkdir(parents=True, exist_ok=True)
        stamp = str(int(time.time() * 1000))
        full_path = debug_dir / f'{prefix}-full-{stamp}.png'
        crop_path = debug_dir / f'{prefix}-crop-{stamp}.png'
        try:
            capture = subprocess.run(
                ['screencapture', '-x', '-R', f'{left},{top},{width},{height}', str(crop_path)],
                capture_output=True,
                text=True,
                timeout=10,
            )
            if capture.returncode != 0:
                raise RuntimeError(capture.stderr.strip() or 'screencapture failed')

            ocr = subprocess.run(
                [self.tesseract_path, str(crop_path), 'stdout', '--psm', '6'],
                capture_output=True,
                text=True,
                timeout=20,
            )
            if ocr.returncode != 0:
                raise RuntimeError(ocr.stderr.strip() or 'tesseract OCR failed')
            return (ocr.stdout or '').strip(), str(crop_path), ''
        except Exception as error:
            return '', str(crop_path), str(error)

    def _read_result_region_ocr(self) -> tuple[str, str, str]:
        return self._read_region_ocr('cy_result_top_left', 'cy_result_bottom_right', 'cy')

    def _extract_estimate_value(self, text: str):
        if not text:
            return None
        patterns = [
            r'Estimate[^0-9]*([0-9]+(?:\.[0-9]+)?)',
            r'\bEstimate\b[\s\S]{0,40}?([0-9]+(?:\.[0-9]+)?)',
            r'\n([0-9]+(?:\.[0-9]+)?)\n(?:90%|80%|Confidence)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                try:
                    return float(match.group(1))
                except ValueError:
                    continue
        return None

    def _extract_confidence_value(self, text: str):
        if not text:
            return None
        patterns = [
            r'Confidence[^0-9]*([0-9]+(?:\.[0-9]+)?)',
            r'\bConfidence\b[\s\S]{0,30}?([0-9]+(?:\.[0-9]+)?)',
            r'Estimate[\s\S]{0,80}?Confidence[^0-9]*([0-9]+(?:\.[0-9]+)?)',
            r'\b([0-9]+(?:\.[0-9]+)?)\b\s*$'
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if match:
                value = match.group(1)
                try:
                    return int(float(value))
                except ValueError:
                    continue
        return None

    def _extract_result_cert_value(self, text: str) -> str:
        if not text:
            return ''
        patterns = [
            r'Cert[^0-9]*([0-9][0-9\s-]{4,})',
            r'Certificate[^0-9]*([0-9][0-9\s-]{4,})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return ''.join(ch for ch in match.group(1) if ch.isdigit())
        return ''

    def _read_matching_result(self, cert_number: str) -> dict:
        timeout = self._timing('result_poll_timeout', 9000)
        interval = max(self._timing('result_poll_interval', 250), 0.1)
        settle_seconds = self._timing('result_settle_wait', 2500)
        stable_required = int(self.layout.get('timing_ms', {}).get('result_stable_polls', 3))
        deadline = time.time() + timeout
        attempts = 0
        first_complete_at = None
        last_complete = None
        last_signature = None
        stable_count = 0
        last = {
            'ocr_text': '',
            'ocr_crop_path': '',
            'ocr_error': '',
            'estimate': None,
            'confidence': None,
            'visible_cert': '',
            'matched': False,
            'attempts': 0,
        }

        while True:
            attempts += 1
            try:
                ocr_text, ocr_crop_path, ocr_error = self._read_result_region_ocr()
                visible_cert = self._extract_result_cert_value(ocr_text)
                estimate = self._extract_estimate_value(ocr_text)
                confidence = self._extract_confidence_value(ocr_text)
                matched = visible_cert == cert_number
                last = {
                    'ocr_text': ocr_text,
                    'ocr_crop_path': ocr_crop_path,
                    'ocr_error': ocr_error,
                    'estimate': estimate if matched else None,
                    'confidence': confidence if matched else None,
                    'visible_cert': visible_cert,
                    'matched': matched,
                    'attempts': attempts,
                }
                if matched and estimate is not None and confidence is not None:
                    signature = (round(float(estimate), 2), int(confidence), visible_cert)
                    if signature == last_signature:
                        stable_count += 1
                    else:
                        stable_count = 1
                        last_signature = signature
                    if first_complete_at is None:
                        first_complete_at = time.time()
                    last_complete = dict(last)
                    last_complete['settle_debug'] = {
                        'first_complete_wait': round(time.time() - first_complete_at, 3),
                        'stable_count': stable_count,
                        'stable_required': stable_required,
                        'settle_seconds': settle_seconds,
                    }
                    if time.time() - first_complete_at >= settle_seconds and stable_count >= stable_required:
                        return last_complete
            except Exception as error:
                last.update({
                    'ocr_text': '',
                    'ocr_crop_path': '',
                    'ocr_error': str(error),
                    'attempts': attempts,
                })

            if time.time() >= deadline:
                return last_complete or last
            time.sleep(interval)

    def submit_cert_lookup(self, cert_number: str, slab_type: str) -> dict:
        timings = {}
        t_submit = time.time()
        last_mark = t_submit

        def mark(name: str):
            nonlocal last_mark
            now = time.time()
            timings[name] = round(now - last_mark, 3)
            last_mark = now

        cert_number = str(cert_number or '').strip()
        if not cert_number:
            raise ValueError('Missing cert_number')

        slab = self._normalize_slab_type(slab_type)
        was_running, window_bounds = self._ensure_app_running()
        mark('ensure_app')
        initial_dismiss_debug = {
            'dismissed': [],
            'dialog_detected': False,
            'dialog_text': '',
            'skipped': not self.needs_dialog_cleanup,
        }
        if self.needs_dialog_cleanup:
            initial_dismiss_debug = self._dismiss_possible_error_dialog()
        mark('initial_dialog_cleanup')
        self._ensure_manual_entry()
        mark('manual_entry')
        scroll_debug = self._select_slab_type(slab)
        mark('select_slab')
        self._activate_app()
        self._click_point('cert_input')
        self._click_point('cert_input')
        self._paste_text(cert_number)
        mark('enter_cert')
        self._click_point('search_button')
        time.sleep(max(self._timing('after_search', 250), 0.35))
        mark('search_wait')

        window_text = ''
        result_read = self._read_matching_result(cert_number)
        estimate = result_read['estimate']
        confidence = result_read['confidence']
        ocr_text = result_read['ocr_text']
        ocr_crop_path = result_read['ocr_crop_path']
        ocr_error = result_read['ocr_error']
        mark('result_ocr')

        dismiss_debug = {
            'dismissed': [],
            'dialog_detected': False,
            'dialog_text': '',
            'skipped': estimate is not None,
        }
        if estimate is None:
            dismiss_debug = self._dismiss_possible_error_dialog()
            self.needs_dialog_cleanup = bool(dismiss_debug.get('dialog_detected')) and not dismiss_debug.get('dismissed')
        else:
            self.needs_dialog_cleanup = False
        mark('post_dialog_cleanup')
        timings['total'] = round(time.time() - t_submit, 3)

        return {
            'cert_number': cert_number,
            'slab_type': slab,
            'submitted': True,
            'status': 'submitted',
            'app_was_running': was_running,
            'window_bounds': window_bounds,
            'cy_buy_price': estimate,
            'cy_confidence': confidence,
            'raw_text': window_text,
            'ocr_text': ocr_text,
            'ocr_crop_path': ocr_crop_path,
            'ocr_error': ocr_error,
            'result_match_debug': {
                'visible_cert': result_read.get('visible_cert', ''),
                'matched': result_read.get('matched', False),
                'attempts': result_read.get('attempts', 0),
            },
            'scroll_debug': scroll_debug,
            'initial_dismiss_debug': initial_dismiss_debug,
            'dismiss_debug': dismiss_debug,
            'timings': timings,
            'message': f'Submitted {slab} cert {cert_number} to {self.app_name}',
            'app_name': self.app_name,
        }

    def get_buy_price(self, cert_number: str) -> dict:
        cert_number = str(cert_number or '').strip()
        if not cert_number:
            raise ValueError('Missing cert_number')

        return self.submit_cert_lookup(cert_number, 'PSA')
