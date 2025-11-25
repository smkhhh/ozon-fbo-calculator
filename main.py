from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify
from flask_login import LoginManager, login_required, current_user
from auth import auth_bp, db, User
from calculator import OzonPriceFinder
import os
import uuid

app = Flask(__name__)
app.secret_key = "your_secret_key"
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///users.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db.init_app(app)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "auth.login"
app.register_blueprint(auth_bp, url_prefix="/auth")

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# Создание таблиц при первом запуске
with app.app_context():
    db.create_all()

@app.route("/")
@login_required
def index():
    return render_template("index.html", username=current_user.username)

# --- Новый API обработчик для загрузки и расчёта ---
@app.route("/api/calculate", methods=["POST"])
@login_required
def api_calculate():
    if "file" not in request.files:
        return jsonify({"success": False, "message": "Файл не выбран"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"success": False, "message": "Имя файла пустое"}), 400

    # Создаём уникальное имя для файла
    unique_id = uuid.uuid4().hex
    temp_dir = "instance"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    temp_excel = os.path.join(temp_dir, f"upload_{unique_id}.xlsx")
    result_excel = os.path.join(temp_dir, f"result_{unique_id}.xlsx")
    file.save(temp_excel)

    target_margin = float(request.form.get("target_margin", 20.0))

    try:
        pf = OzonPriceFinder()
        pf.calculate_file(temp_excel, target_margin_pct=target_margin, output_excel=result_excel)
        # Возвращаем имя итогового файла для скачивания
        return jsonify({"success": True, "download_url": url_for("download_result", filename=f"result_{unique_id}.xlsx")})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@app.route("/api/download/<filename>")
@login_required
def download_result(filename):
    temp_dir = "instance"
    file_path = os.path.join(temp_dir, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "Файл не найден", 404

# --- Страница для ручного расчёта (если будет нужна) ---
@app.route("/calculate", methods=["GET", "POST"])
@login_required
def calculate():
    if request.method == "POST":
        # Старая реализация под загрузку по локальному пути, для совместимости
        target_margin = float(request.form.get("target_margin", 20))
        excel_path = request.form.get("excel_path")
        pf = OzonPriceFinder()
        try:
            result_df = pf.calculate_file(excel_path, target_margin_pct=target_margin)
            # Это реализовано если нужен просмотр как HTML-таблица:
            return render_template(
                "index.html",  # index.html, тк results.html у вас нет
                tables=[result_df.to_html(classes="data")],
                titles=result_df.columns.values,
                username=current_user.username
            )
        except Exception as e:
            return f"Ошибка: {e}", 500
    return render_template("index.html", username=current_user.username)

if __name__ == "__main__":
    app.run(debug=True)
