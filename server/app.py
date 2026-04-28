"""
SYMMETRY Lab 予約・決済サーバー（本番対応版）
FastAPI + Stripe Checkout + SQLite + Excel出力 + メール通知
"""

import json
import os
import smtplib
import sqlite3
import traceback
import urllib.request
import urllib.error
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
from pathlib import Path
from typing import Optional

import openpyxl
import stripe
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl.styles import Alignment, Font, PatternFill
from pydantic import BaseModel

load_dotenv()

stripe.api_key = os.getenv("STRIPE_SECRET_KEY")
BASE_URL = os.getenv("BASE_URL", "http://localhost:8000")
ADMIN_KEY = os.getenv("ADMIN_KEY", "symmetry-admin-2026")
DB_PATH = os.getenv("DB_PATH", "bookings.db")

# DBパスの親ディレクトリが書き込み不可ならローカル bookings.db にフォールバック
# （Render Freeプランで永続ディスク /var/data が使えないケース等）
_db_dir = os.path.dirname(DB_PATH)
if _db_dir:
    try:
        os.makedirs(_db_dir, exist_ok=True)
        _probe = os.path.join(_db_dir, ".write_probe")
        with open(_probe, "w") as _f:
            _f.write("")
        os.remove(_probe)
    except (PermissionError, OSError) as _e:
        print(f"[起動] DB_PATH={DB_PATH} の親ディレクトリが書き込み不可（{_e}）→ ./bookings.db にフォールバック")
        DB_PATH = "bookings.db"

DEFAULT_TRAINING_DATES = Path(__file__).parent / "training_dates.json"
TRAINING_DATES_PATH = Path(os.getenv("TRAINING_DATES_PATH", str(DEFAULT_TRAINING_DATES)))
DEFAULT_REFERRAL_CODES = Path(__file__).parent / "referral_codes.json"
REFERRAL_CODES_PATH = Path(os.getenv("REFERRAL_CODES_PATH", str(DEFAULT_REFERRAL_CODES)))

# 永続ディスク側にファイルが無ければ、初回起動時にリポジトリ同梱のデフォルトをコピー
# （失敗しても起動は継続する — 書き込み不可な環境でも最低限デフォルトをロードして動作させる）
try:
    if not TRAINING_DATES_PATH.exists() and DEFAULT_TRAINING_DATES.exists():
        TRAINING_DATES_PATH.parent.mkdir(parents=True, exist_ok=True)
        TRAINING_DATES_PATH.write_text(DEFAULT_TRAINING_DATES.read_text(encoding="utf-8"), encoding="utf-8")
except Exception as _e:
    print(f"[起動] training_dates.json の永続化領域への初期化に失敗（デフォルトを使用）: {_e}")
    TRAINING_DATES_PATH = DEFAULT_TRAINING_DATES

try:
    if not REFERRAL_CODES_PATH.exists() and DEFAULT_REFERRAL_CODES.exists():
        REFERRAL_CODES_PATH.parent.mkdir(parents=True, exist_ok=True)
        REFERRAL_CODES_PATH.write_text(DEFAULT_REFERRAL_CODES.read_text(encoding="utf-8"), encoding="utf-8")
except Exception as _e:
    print(f"[起動] referral_codes.json の永続化領域への初期化に失敗（デフォルトを使用）: {_e}")
    REFERRAL_CODES_PATH = DEFAULT_REFERRAL_CODES

# --- メール設定 ---
SMTP_EMAIL = os.getenv("SMTP_EMAIL", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL", SMTP_EMAIL)

# --- LINE設定 ---
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")

app = FastAPI(title="SYMMETRY Lab Booking API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# --- データモデル ---
class CheckoutRequest(BaseModel):
    training_type: str
    training_date: str
    customer_name: str
    customer_email: str
    customer_phone: str = ""
    customer_company: str = ""
    sessions: int = 1
    booking_notes: str = ""  # 第1〜第3希望など補足情報（Stripe metadataへ退避）
    referral_code: str = ""


# --- SQLite ---
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            booking_id TEXT,
            created_at TEXT,
            training_type TEXT,
            training_name TEXT,
            training_date TEXT,
            customer_name TEXT,
            customer_email TEXT,
            customer_phone TEXT,
            customer_company TEXT,
            amount INTEGER,
            payment_status TEXT DEFAULT 'paid',
            stripe_session_id TEXT UNIQUE,
            notes TEXT DEFAULT ''
        )
    """)
    conn.execute("UPDATE bookings SET payment_status = 'paid' WHERE payment_status = '完了'")
    conn.execute("UPDATE bookings SET created_at = REPLACE(created_at, '/', '-') WHERE created_at LIKE '____/__/__%'")
    conn.commit()
    conn.close()


def save_booking(session_data: dict) -> bool:
    """予約をDBに保存。新規挿入時True、既存（重複）ならFalseを返す"""
    metadata = session_data.get("metadata", {})
    session_id = session_data.get("id", "")
    short_id = session_id[-8:] if session_id else ""
    amount = session_data.get("amount_total", 0)
    if amount and isinstance(amount, int) and amount > 1000:
        pass  # already in yen
    else:
        amount = int(metadata.get("price", 0))

    conn = get_db()
    try:
        conn.execute("""
            INSERT OR IGNORE INTO bookings
            (booking_id, created_at, training_type, training_name, training_date,
             customer_name, customer_email, customer_phone, customer_company,
             amount, payment_status, stripe_session_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'paid', ?)
        """, (
            short_id,
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            metadata.get("training_type", ""),
            metadata.get("training_name", ""),
            metadata.get("training_date", ""),
            metadata.get("customer_name", ""),
            metadata.get("customer_email", ""),
            metadata.get("customer_phone", ""),
            metadata.get("customer_company", ""),
            amount,
            session_id,
        ))
        conn.commit()
        inserted = conn.total_changes > 0
        if inserted:
            print(f"[予約保存] {metadata.get('customer_name', '')} - {metadata.get('training_name', '')} ({metadata.get('training_date', '')})")
        else:
            print(f"[予約スキップ] session_id={session_id} は既に記録済み")
        return inserted
    finally:
        conn.close()


def resolve_training(training_type: str, data: Optional[dict] = None):
    """case_interview_new/mid → case_interview へのフォールバックを一元化"""
    if data is None:
        data = load_training_dates()
    training = data.get(training_type)
    if not training and training_type in ("case_interview_new", "case_interview_mid"):
        training = data.get("case_interview")
    return training


def count_bookings_for_date(training_type: str, date: str) -> int:
    training = resolve_training(training_type)
    type_name = training.get("name", "") if training else ""
    conn = get_db()
    cursor = conn.execute(
        "SELECT COUNT(*) FROM bookings WHERE training_name = ? AND training_date = ?",
        (type_name, date)
    )
    count = cursor.fetchone()[0]
    conn.close()
    return count


def load_training_dates():
    with open(TRAINING_DATES_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


# --- 紹介コード ---
def load_referral_codes() -> dict:
    """紹介コード一覧を読み込む（無ければ空構造）"""
    try:
        with open(REFERRAL_CODES_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {"codes": []}


def save_referral_codes(data: dict):
    """紹介コード一覧を保存"""
    try:
        REFERRAL_CODES_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(REFERRAL_CODES_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[紹介コード] 保存失敗: {e}")


def find_referral_code(code: str) -> Optional[dict]:
    """大文字小文字無視でコードを検索（戻り値は元データへの参照）"""
    if not code:
        return None
    code_norm = code.strip().upper()
    data = load_referral_codes()
    for entry in data.get("codes", []):
        if entry.get("code", "").strip().upper() == code_norm:
            return entry
    return None


def validate_referral_code(code: str, training_type: str = "") -> dict:
    """
    紹介コードを検証し、適用情報を返す。
    戻り値: {"valid": bool, "reason": str, "discount_type": "rate"|"amount", "discount_value": ..., "label": "..."}
    """
    entry = find_referral_code(code)
    if not entry:
        return {"valid": False, "reason": "コードが見つかりません"}
    if not entry.get("active", True):
        return {"valid": False, "reason": "現在停止中のコードです"}

    # 期限チェック
    expires = entry.get("expires", "")
    if expires:
        try:
            today = datetime.now().date()
            exp_date = datetime.strptime(expires, "%Y-%m-%d").date()
            if today > exp_date:
                return {"valid": False, "reason": f"有効期限切れ（{expires}まで）"}
        except ValueError:
            pass

    # 利用上限チェック
    max_uses = entry.get("max_uses")
    used_count = entry.get("used_count", 0)
    if max_uses is not None and used_count >= max_uses:
        return {"valid": False, "reason": "利用上限に達しています"}

    # 対象研修種別チェック
    applies_to = entry.get("applies_to", []) or []
    if applies_to and training_type:
        # case_interview_new/mid → case_interview の正規化を反映
        normalized = "case_interview" if training_type in ("case_interview_new", "case_interview_mid") else training_type
        if normalized not in applies_to and training_type not in applies_to:
            return {"valid": False, "reason": "対象外の研修種別です"}

    discount_type = entry.get("discount_type", "amount")
    discount_value = entry.get("discount_value", 0)

    if discount_type == "rate":
        label = f"{int(discount_value * 100)}% OFF"
    else:
        label = f"¥{int(discount_value):,} 割引"

    return {
        "valid": True,
        "reason": "適用可能",
        "code": entry.get("code", ""),
        "discount_type": discount_type,
        "discount_value": discount_value,
        "label": label,
        "note": entry.get("note", ""),
    }


def calc_discounted_total(original_total: int, validation: dict) -> int:
    """検証済みコード情報から割引後の合計金額を計算"""
    if not validation.get("valid"):
        return original_total
    dtype = validation.get("discount_type")
    dval = validation.get("discount_value", 0)
    if dtype == "rate":
        discounted = int(round(original_total * (1 - float(dval))))
    elif dtype == "amount":
        discounted = int(original_total - int(dval))
    else:
        discounted = original_total
    # 0円以下は0円に丸める（Stripeは50円未満決済不可なので、最低50円に）
    return max(50, discounted)


def increment_referral_use(code: str):
    """紹介コードの利用回数を+1"""
    if not code:
        return
    code_norm = code.strip().upper()
    data = load_referral_codes()
    changed = False
    for entry in data.get("codes", []):
        if entry.get("code", "").strip().upper() == code_norm:
            entry["used_count"] = entry.get("used_count", 0) + 1
            changed = True
            break
    if changed:
        save_referral_codes(data)
        print(f"[紹介コード] 利用カウント+1: {code_norm}")


# --- メール送信 ---
def send_email(to_email: str, subject: str, html_body: str):
    """Gmail SMTP でメール送信"""
    if not SMTP_EMAIL or not SMTP_PASSWORD:
        print(f"[メール] SMTP未設定のためスキップ: {subject} → {to_email}")
        return False
    try:
        msg = MIMEMultipart("alternative")
        msg["From"] = f"SYMMETRY Lab <{SMTP_EMAIL}>"
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(html_body, "html", "utf-8"))

        smtp_host = os.getenv("SMTP_HOST", "smtp.office365.com")
        smtp_port = int(os.getenv("SMTP_PORT", "587"))

        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, to_email, msg.as_string())

        print(f"[メール送信成功] {subject} → {to_email}")
        return True
    except Exception as e:
        print(f"[メール送信失敗] {subject} → {to_email}: {e}")
        traceback.print_exc()
        return False


def send_booking_confirmation(metadata: dict, amount: int):
    """予約確認メールを顧客に送信"""
    customer_name = metadata.get("customer_name", "")
    customer_email = metadata.get("customer_email", "")
    training_name = metadata.get("training_name", "")
    training_date = metadata.get("training_date", "")
    sessions = metadata.get("sessions", "1")

    if not customer_email:
        return

    subject = f"【SYMMETRY Lab】{training_name} お申込み確認"

    html = f"""
    <div style="max-width:600px;margin:0 auto;font-family:'Helvetica Neue',Arial,sans-serif;color:#1F2937;line-height:1.8;">
      <div style="border-top:3px solid #0ABAB5;padding:40px 32px;">
        <h1 style="font-size:20px;font-weight:700;color:#1F2937;margin:0 0 8px;">
          お申込みありがとうございます
        </h1>
        <p style="font-size:13px;color:#6B7280;margin:0 0 32px;">
          SYMMETRY Lab株式会社
        </p>

        <p style="font-size:15px;">
          {customer_name} 様<br><br>
          この度は <strong>{training_name}</strong> にお申込みいただき、誠にありがとうございます。<br>
          以下の内容でご予約を承りました。
        </p>

        <div style="background:#F9FAFB;border-left:3px solid #0ABAB5;padding:20px 24px;margin:28px 0;border-radius:2px;">
          <table style="width:100%;font-size:14px;border-collapse:collapse;">
            <tr>
              <td style="padding:8px 0;color:#6B7280;width:140px;">プログラム</td>
              <td style="padding:8px 0;font-weight:600;">{training_name}</td>
            </tr>
            <tr>
              <td style="padding:8px 0;color:#6B7280;">日程</td>
              <td style="padding:8px 0;font-weight:600;">{training_date}</td>
            </tr>
            <tr>
              <td style="padding:8px 0;color:#6B7280;">セッション数</td>
              <td style="padding:8px 0;font-weight:600;">{sessions}回</td>
            </tr>
            <tr style="border-top:1px solid #E5E7EB;">
              <td style="padding:12px 0 8px;color:#6B7280;">お支払い金額</td>
              <td style="padding:12px 0 8px;font-weight:700;font-size:18px;color:#0ABAB5;">
                ¥{amount:,}
              </td>
            </tr>
          </table>
        </div>

        <h2 style="font-size:15px;font-weight:700;color:#1F2937;margin:32px 0 12px;">
          今後の流れ
        </h2>
        <ol style="font-size:14px;padding-left:20px;color:#4B5563;">
          <li style="margin-bottom:8px;">本メールで予約内容をご確認ください</li>
          <li style="margin-bottom:8px;">セッション当日、Zoomリンクをメールでお送りします</li>
          <li style="margin-bottom:8px;">お時間になりましたらZoomにてご参加ください</li>
        </ol>

        <p style="font-size:14px;color:#4B5563;margin-top:28px;">
          ご不明点がございましたら、お気軽にご連絡ください。<br>
          お会いできることを楽しみにしております。
        </p>

        <div style="border-top:1px solid #E5E7EB;margin-top:40px;padding-top:20px;">
          <p style="font-size:12px;color:#9CA3AF;margin:0;">
            SYMMETRY Lab株式会社<br>
            Email: {SMTP_EMAIL}<br>
            Web: https://symmetrylab.jp
          </p>
        </div>
      </div>
    </div>
    """

    send_email(customer_email, subject, html)

    # 管理者にも通知
    admin_subject = f"[新規予約] {customer_name}様 - {training_name} ({training_date})"
    admin_html = f"""
    <div style="font-family:sans-serif;font-size:14px;color:#333;line-height:1.8;">
      <h2 style="color:#0ABAB5;">新規予約通知</h2>
      <table style="border-collapse:collapse;">
        <tr><td style="padding:4px 16px 4px 0;color:#888;">氏名</td><td><strong>{customer_name}</strong></td></tr>
        <tr><td style="padding:4px 16px 4px 0;color:#888;">メール</td><td>{customer_email}</td></tr>
        <tr><td style="padding:4px 16px 4px 0;color:#888;">電話</td><td>{metadata.get('customer_phone', '')}</td></tr>
        <tr><td style="padding:4px 16px 4px 0;color:#888;">プログラム</td><td>{training_name}</td></tr>
        <tr><td style="padding:4px 16px 4px 0;color:#888;">日程</td><td>{training_date}</td></tr>
        <tr><td style="padding:4px 16px 4px 0;color:#888;">セッション数</td><td>{sessions}回</td></tr>
        <tr><td style="padding:4px 16px 4px 0;color:#888;">金額</td><td><strong>¥{amount:,}</strong></td></tr>
      </table>
    </div>
    """
    if ADMIN_EMAIL:
        send_email(ADMIN_EMAIL, admin_subject, admin_html)


# --- LINE Messaging API ---
def send_line_push(user_id: str, messages: list) -> bool:
    """LINE Push API でメッセージ送信"""
    if not LINE_CHANNEL_ACCESS_TOKEN or not user_id:
        return False
    try:
        req = urllib.request.Request(
            "https://api.line.me/v2/bot/message/push",
            method="POST",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}",
            },
            data=json.dumps({"to": user_id, "messages": messages}).encode("utf-8"),
        )
        with urllib.request.urlopen(req, timeout=10) as res:
            res.read()
        print(f"[LINE送信成功] → {user_id}")
        return True
    except Exception as e:
        print(f"[LINE送信失敗] {e}")
        traceback.print_exc()
        return False


def send_line_booking_notification(metadata: dict, amount: int):
    """予約完了後、LINEトークに確認メッセージ送信"""
    company = metadata.get("customer_company", "")
    if not company.startswith("LINE:"):
        return
    user_id = company.replace("LINE:", "").strip()
    if not user_id:
        return

    customer_name = metadata.get("customer_name", "")
    training_name = metadata.get("training_name", "")
    training_date = metadata.get("training_date", "")
    sessions = metadata.get("sessions", "1")

    bubble = {
        "type": "bubble",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": "#0ABAB5",
            "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": "予約完了", "color": "#FFFFFF", "weight": "bold", "size": "lg"},
                {"type": "text", "text": "SYMMETRY Lab", "color": "#E6F7F6", "size": "xs", "margin": "sm"},
            ],
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": f"{customer_name} 様", "weight": "bold", "size": "md"},
                {"type": "text", "text": "お申込みありがとうございます。", "size": "sm", "color": "#6B7280", "margin": "sm", "wrap": True},
                {"type": "separator", "margin": "lg"},
                {
                    "type": "box", "layout": "vertical", "margin": "lg", "spacing": "sm",
                    "contents": [
                        {"type": "box", "layout": "baseline", "contents": [
                            {"type": "text", "text": "プログラム", "size": "xs", "color": "#6B7280", "flex": 3},
                            {"type": "text", "text": training_name, "size": "xs", "flex": 5, "wrap": True, "weight": "bold"},
                        ]},
                        {"type": "box", "layout": "baseline", "contents": [
                            {"type": "text", "text": "希望日時", "size": "xs", "color": "#6B7280", "flex": 3},
                            {"type": "text", "text": training_date, "size": "xs", "flex": 5, "wrap": True, "weight": "bold"},
                        ]},
                        {"type": "box", "layout": "baseline", "contents": [
                            {"type": "text", "text": "セッション数", "size": "xs", "color": "#6B7280", "flex": 3},
                            {"type": "text", "text": f"{sessions}回", "size": "xs", "flex": 5, "weight": "bold"},
                        ]},
                        {"type": "box", "layout": "baseline", "contents": [
                            {"type": "text", "text": "お支払い", "size": "xs", "color": "#6B7280", "flex": 3},
                            {"type": "text", "text": f"¥{amount:,}", "size": "sm", "flex": 5, "weight": "bold", "color": "#0ABAB5"},
                        ]},
                    ],
                },
                {"type": "separator", "margin": "lg"},
                {"type": "text", "text": "担当者より24時間以内にご連絡いたします。", "size": "xs", "color": "#6B7280", "margin": "lg", "wrap": True},
            ],
        },
    }

    messages = [
        {"type": "flex", "altText": f"{training_name} お申込み完了", "contents": bubble},
    ]
    send_line_push(user_id, messages)


# --- APIエンドポイント ---
@app.get("/api/available-dates")
async def get_available_dates(training_type: str, date: str = ""):
    training = resolve_training(training_type)
    if not training:
        raise HTTPException(status_code=404, detail="研修種別が見つかりません")

    avail_slots = training.get("available_slots", {})
    time_slots = training.get("time_slots", [])

    if date:
        if date < datetime.now().strftime("%Y-%m-%d"):
            return {"time_slots": []}
        # available_slotsが空なら全スロット許可、設定済みならその日のスロットのみ
        if avail_slots:
            allowed_times = avail_slots.get(date, [])
            if not allowed_times:
                return {"time_slots": []}
        else:
            allowed_times = time_slots

        available = []
        for slot in allowed_times:
            booked = count_bookings_for_date(training_type, f"{date} {slot}")
            remaining = training["max_capacity"] - booked
            if remaining > 0:
                available.append({"time": slot, "slots_remaining": remaining})
        return {"time_slots": available}

    # 日付一覧: available_slotsのキー
    available_dates = list(avail_slots.keys()) if avail_slots else []

    return {
        "training_name": training["name"],
        "price": training["price"],
        "price_label": training["price_label"],
        "time_slots": time_slots,
        "available_dates": available_dates,
        "available_slots": avail_slots,
    }


@app.post("/api/create-checkout-session")
async def create_checkout_session(req: CheckoutRequest):
    # 入力の基本バリデーション（サーバ側でのリクエスト内容を必ずログに出す）
    print(f"[checkout] type={req.training_type} date={req.training_date!r} email={req.customer_email} sessions={req.sessions} notes_len={len(req.booking_notes)}")

    if not stripe.api_key:
        raise HTTPException(status_code=400, detail="サーバ設定エラー: 決済サービスが未設定です。管理者へお問い合わせください。")
    if not BASE_URL or not BASE_URL.startswith(("http://", "https://")):
        raise HTTPException(status_code=400, detail="サーバ設定エラー: BASE_URLが正しく設定されていません。")
    if not req.customer_email or "@" not in req.customer_email:
        raise HTTPException(status_code=400, detail="メールアドレスの形式が正しくありません。")
    if not req.customer_name.strip():
        raise HTTPException(status_code=400, detail="お名前が空です。")

    training = resolve_training(req.training_type)
    if not training:
        raise HTTPException(status_code=400, detail=f"無効な研修種別です: {req.training_type}")

    avail_slots = training.get("available_slots", {})
    # training_date は "2026-04-25 14:00" 形式を前提。先頭の日付のみ抽出
    date_part = req.training_date.split(" ")[0] if " " in req.training_date else req.training_date
    if avail_slots and date_part not in avail_slots:
        raise HTTPException(status_code=400, detail=f"この日程は予約可能日として設定されていません（{date_part}）")

    booked = count_bookings_for_date(req.training_type, req.training_date)
    if booked >= training["max_capacity"]:
        raise HTTPException(status_code=400, detail="この日程は定員に達しています")

    qty = max(1, int(req.sessions))

    # ケース面接対策はパッケージ価格（割引適用）。それ以外は単価×数量
    CASE_PACKAGE_PRICES = {1: 7000, 2: 13580, 3: 19950, 5: 32200, 10: 63000}
    if req.training_type in ("case_interview", "case_interview_new", "case_interview_mid") and qty in CASE_PACKAGE_PRICES:
        original_total = CASE_PACKAGE_PRICES[qty]
    else:
        original_total = training["price"] * qty

    # 紹介コード検証＆割引適用
    referral_validation = {"valid": False}
    discount_amount = 0
    if req.referral_code:
        referral_validation = validate_referral_code(req.referral_code, req.training_type)
        if referral_validation.get("valid"):
            discounted_total = calc_discounted_total(original_total, referral_validation)
            discount_amount = original_total - discounted_total
            total_price = discounted_total
        else:
            # 無効コードは入力ミスの可能性がある。エラーで止めずログだけ残して通常価格で進める
            print(f"[checkout] 紹介コード無効: {req.referral_code} → {referral_validation.get('reason')}")
            total_price = original_total
    else:
        total_price = original_total

    # Stripeのline_itemsは「単価×数量」モデルなので、合計額を unit_amount に入れて quantity=1 で渡す
    unit_amount = total_price
    stripe_quantity = 1

    # Stripe product name は250文字制限。日時部分だけを載せ、希望一覧は metadata へ
    name_with_qty = f"{training['name']} - {req.training_date}"
    if qty > 1:
        name_with_qty += f"（{qty}セッションパッケージ）"
    if len(name_with_qty) > 240:
        name_with_qty = name_with_qty[:237] + "..."

    # Stripe metadataは各value 500文字制限。booking_notesは念のため切り詰め
    notes = (req.booking_notes or "")[:490]

    try:
        session = stripe.checkout.Session.create(
            payment_method_types=["card"],
            line_items=[{
                "price_data": {
                    "currency": "jpy",
                    "product_data": {
                        "name": name_with_qty,
                        "description": f"SYMMETRY Lab {training['name']}"
                    },
                    "unit_amount": unit_amount,
                },
                "quantity": stripe_quantity,
            }],
            mode="payment",
            success_url=f"{BASE_URL}/booking.html?success=true&session_id={{CHECKOUT_SESSION_ID}}",
            cancel_url=f"{BASE_URL}/lp-case.html?canceled=true",
            customer_email=req.customer_email,
            metadata={
                "training_type": req.training_type,
                "training_name": training["name"],
                "training_date": req.training_date,
                "customer_name": req.customer_name,
                "customer_email": req.customer_email,
                "customer_phone": req.customer_phone,
                "customer_company": req.customer_company,
                "price": str(total_price),
                "original_price": str(original_total),
                "discount_amount": str(discount_amount),
                "referral_code": req.referral_code if referral_validation.get("valid") else "",
                "sessions": str(qty),
                "booking_notes": notes,
            }
        )
        print(f"[checkout] OK session={session.id}")
        return {"checkout_url": session.url}
    except stripe.error.StripeError as e:
        print(f"[checkout] StripeError: {type(e).__name__}: {e}")
        traceback.print_exc()
        # Stripeエラーは設定/入力起因が多いので400で返してUI側で文言を出せるように
        raise HTTPException(status_code=400, detail=f"決済セッションの作成に失敗しました: {str(e)}")
    except Exception as e:
        print(f"[checkout] Unexpected: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"予期しないエラーが発生しました: {str(e)}")


@app.get("/api/confirm-booking")
async def confirm_booking(session_id: str):
    try:
        session = stripe.checkout.Session.retrieve(session_id)
        if session.payment_status != "paid":
            raise HTTPException(status_code=400, detail="決済が完了していません")

        md = session.metadata
        session_data = {
            "id": session.id,
            "amount_total": session.amount_total,
            "metadata": {
                "training_type": md["training_type"] if "training_type" in md else "",
                "training_name": md["training_name"] if "training_name" in md else "",
                "training_date": md["training_date"] if "training_date" in md else "",
                "customer_name": md["customer_name"] if "customer_name" in md else "",
                "customer_email": md["customer_email"] if "customer_email" in md else "",
                "customer_phone": md["customer_phone"] if "customer_phone" in md else "",
                "customer_company": md["customer_company"] if "customer_company" in md else "",
                "price": md["price"] if "price" in md else "0",
            }
        }
        inserted = save_booking(session_data)
        print(f"[予約確認] session_id={session_id} inserted={inserted}")

        # 新規予約のときのみ通知送信＋紹介コード使用カウント（confirm-booking多重呼び出し対策）
        if inserted:
            amount = session_data.get("amount_total", 0)
            send_booking_confirmation(session_data["metadata"], amount)
            send_line_booking_notification(session_data["metadata"], amount)
            # 紹介コード使用回数を+1
            used_code = session_data["metadata"].get("referral_code", "")
            if used_code:
                increment_referral_use(used_code)
            return {"status": "ok", "message": "予約を記録しました"}
        else:
            return {"status": "ok", "message": "予約は既に記録済みです", "already_recorded": True}
    except stripe.error.StripeError as e:
        raise HTTPException(status_code=400, detail=f"セッション情報の取得に失敗: {str(e)}")


@app.get("/api/bookings/export")
async def export_bookings(key: str = ""):
    """予約一覧をExcelでダウンロード"""
    if key != ADMIN_KEY:
        raise HTTPException(status_code=403, detail="認証が必要です")

    try:
        conn = get_db()
        rows = conn.execute("SELECT * FROM bookings ORDER BY id DESC").fetchall()
        conn.close()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "bookings"

        headers = [
            "予約ID", "申込日時", "研修種別", "研修日",
            "氏名", "メールアドレス", "電話番号", "会社名",
            "金額", "決済ステータス", "Stripe Session ID", "備考"
        ]
        ws.append(headers)

        for row in rows:
            ws.append([
                row["booking_id"], row["created_at"], row["training_name"],
                row["training_date"], row["customer_name"], row["customer_email"],
                row["customer_phone"], row["customer_company"], row["amount"],
                row["payment_status"], row["stripe_session_id"], row["notes"]
            ])

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=bookings.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/bookings")
async def list_bookings(key: str = ""):
    """予約一覧をJSON形式で取得"""
    if key != ADMIN_KEY:
        raise HTTPException(status_code=403, detail="認証が必要です")
    conn = get_db()
    rows = conn.execute("SELECT * FROM bookings ORDER BY id DESC").fetchall()
    conn.close()
    return [dict(row) for row in rows]


@app.get("/api/admin/blocked-dates")
async def get_blocked_dates(key: str = ""):
    """予約可能日時一覧を取得"""
    if key != ADMIN_KEY:
        raise HTTPException(status_code=403, detail="認証が必要です")
    data = load_training_dates()
    result = {}
    for t_type, t_data in data.items():
        result[t_type] = {
            "name": t_data["name"],
            "time_slots": t_data.get("time_slots", []),
            "available_slots": t_data.get("available_slots", {}),
        }
    return result


@app.post("/api/admin/blocked-dates")
async def update_blocked_dates(request: Request, key: str = ""):
    """予約可能日時を更新"""
    if key != ADMIN_KEY:
        raise HTTPException(status_code=403, detail="認証が必要です")
    body = await request.json()
    training_type = body.get("training_type", "")
    available_slots = body.get("available_slots", None)

    data = load_training_dates()
    if training_type not in data:
        raise HTTPException(status_code=400, detail="無効な研修種別です")

    if available_slots is not None:
        data[training_type]["available_slots"] = available_slots
    with open(TRAINING_DATES_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return {"status": "ok"}


@app.get("/api/admin/stats")
async def get_stats(key: str = ""):
    """ダッシュボード統計"""
    if key != ADMIN_KEY:
        raise HTTPException(status_code=403, detail="認証が必要です")
    conn = get_db()
    total = conn.execute("SELECT COUNT(*) FROM bookings").fetchone()[0]
    total_revenue = conn.execute("SELECT COALESCE(SUM(amount), 0) FROM bookings WHERE payment_status = 'paid'").fetchone()[0]
    this_month = datetime.now().strftime("%Y-%m")
    monthly = conn.execute("SELECT COUNT(*) FROM bookings WHERE created_at LIKE ?", (f"{this_month}%",)).fetchone()[0]
    monthly_revenue = conn.execute("SELECT COALESCE(SUM(amount), 0) FROM bookings WHERE payment_status = 'paid' AND created_at LIKE ?", (f"{this_month}%",)).fetchone()[0]
    conn.close()
    return {
        "total_bookings": total,
        "total_revenue": total_revenue,
        "monthly_bookings": monthly,
        "monthly_revenue": monthly_revenue,
    }


@app.get("/api/validate-referral")
async def api_validate_referral(code: str = "", training_type: str = ""):
    """紹介コードのリアルタイム検証エンドポイント"""
    return validate_referral_code(code, training_type)


@app.get("/api/admin/referral-codes")
async def admin_get_referral_codes(key: str = ""):
    """全紹介コード一覧（管理画面用）"""
    if key != ADMIN_KEY:
        raise HTTPException(status_code=403, detail="認証が必要です")
    return load_referral_codes()


@app.post("/api/admin/referral-codes")
async def admin_save_referral_codes(request: Request, key: str = ""):
    """紹介コード一覧を上書き保存（管理画面用）"""
    if key != ADMIN_KEY:
        raise HTTPException(status_code=403, detail="認証が必要です")
    body = await request.json()
    if "codes" not in body or not isinstance(body["codes"], list):
        raise HTTPException(status_code=400, detail="codesリストが必要です")
    save_referral_codes({"codes": body["codes"]})
    return {"status": "ok", "count": len(body["codes"])}


@app.get("/api/health")
async def health():
    return {"status": "ok", "time": datetime.now().isoformat()}


@app.on_event("startup")
async def startup():
    init_db()
    print(f"[起動] SYMMETRY Lab 予約サーバー - {BASE_URL}")


# 静的ファイル配信（最後にマウント）
static_dir = os.getenv("WEBSITE_DIR", os.path.join(os.path.dirname(__file__), ".."))
if os.path.exists(static_dir):
    app.mount("/", StaticFiles(directory=static_dir, html=True), name="static")
