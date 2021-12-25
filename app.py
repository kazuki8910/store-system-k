from flask import Flask
app = Flask(__name__)

@app.route('/')
def root():
    # import func
    return "完了"
 
if __name__ == '__main__':
    app.run()