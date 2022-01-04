from rq import Queue
from worker import conn
from bottle import route, run


q = Queue(connection=conn)

@route('/index')
def index():
    result = q.enqueue(background_process, '引数１')
    return result

def background_process(name):
    import func
    return name * 10


run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))