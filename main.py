from flask import Flask, render_template, request
from flask_login import LoginManager, login_required, current_user
from auth import auth_bp, db, User
from calculator import OzonPriceFinder

app = Flask(__name__)
app.secret_key = "your_secret_key"
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///users.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db.init_app(app)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'auth.login'

app.register_blueprint(auth_bp, url_prefix='/auth')

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

@app.before_first_request
def create_tables():
    db.create_all()

@app.route('/')
@login_required
def index():
    return f"Добро пожаловать, {current_user.username}!"

@app.route('/calculate', methods=['GET', 'POST'])
@login_required
def calculate():
    if request.method == 'POST':
        target_margin = float(request.form.get('target_margin', 20))
        excel_path = request.form.get('excel_path')
        pf = OzonPriceFinder()
        try:
            result_df = pf.calculate_file(excel_path, target_margin_pct=target_margin)
            return render_template('results.html', tables=[result_df.to_html(classes='data')], titles=result_df.columns.values)
        except Exception as e:
            return f"Ошибка: {e}"
    return render_template('calculate.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
