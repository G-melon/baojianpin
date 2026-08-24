import traceback

log = open("run_health_8899.pythonw.log", "a", encoding="utf-8")
try:
    import sys

    sys.stdout = log
    sys.stderr = log

    import server

    server.app.run(host="0.0.0.0", port=8899, debug=False)
except BaseException:
    traceback.print_exc(file=log)
    log.flush()
    raise
