#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BACKEND_DIR="$ROOT_DIR/backend"
FRONTEND_DIR="$ROOT_DIR/frontend"
SESSION_NAME="${HANDWRITING_TMUX_SESSION:-handwriting-web}"
BACKEND_PORT=5005
FRONTEND_PORT=8080
BACKEND_PYTHON="${BACKEND_PYTHON:-$BACKEND_DIR/.venv/bin/python}"

usage() {
  cat <<EOF
Usage:
  ./start-dev.sh           Start or restart backend and frontend
  ./start-dev.sh --attach  Start or restart, then attach tmux
  ./start-dev.sh --stop    Stop the handwriting-web tmux session

Frontend: http://localhost:${FRONTEND_PORT}/
Backend:  http://127.0.0.1:${BACKEND_PORT}
tmux:     ${SESSION_NAME}
EOF
}

if [[ "${1:-}" == "--help" || "${1:-}" == "-h" ]]; then
  usage
  exit 0
fi

if [[ "${1:-}" == "--stop" ]]; then
  if tmux has-session -t "$SESSION_NAME" 2>/dev/null; then
    tmux kill-session -t "$SESSION_NAME"
    echo "Stopped tmux session: $SESSION_NAME"
  else
    echo "No tmux session to stop: $SESSION_NAME"
  fi
  exit 0
fi

if [[ "${1:-}" != "" && "${1:-}" != "--attach" ]]; then
  usage >&2
  exit 1
fi

if ! command -v tmux >/dev/null 2>&1; then
  echo "tmux is required. Install tmux or add it to PATH, then rerun this script." >&2
  exit 1
fi

if ! command -v lsof >/dev/null 2>&1; then
  echo "lsof is required to check whether ports are already occupied." >&2
  exit 1
fi

if [[ ! -x "$BACKEND_PYTHON" ]]; then
  echo "Backend Python not found or not executable: $BACKEND_PYTHON" >&2
  echo "Create the backend environment first, then rerun this script." >&2
  exit 1
fi

if [[ ! -d "$FRONTEND_DIR/node_modules" ]]; then
  echo "frontend/node_modules is missing. Run 'npm install' in frontend first." >&2
  exit 1
fi

port_listener() {
  local port="$1"
  lsof -nP -iTCP:"$port" -sTCP:LISTEN 2>/dev/null || true
}

ensure_port_free() {
  local port="$1"
  local name="$2"
  local listener
  listener="$(port_listener "$port")"
  if [[ -n "$listener" ]]; then
    echo "Port $port for $name is already occupied by another process:" >&2
    echo "$listener" >&2
    echo "Stop that process first, then rerun ./start-dev.sh." >&2
    exit 1
  fi
}

wait_for_port() {
  local port="$1"
  local name="$2"
  local window="$3"

  for _ in {1..30}; do
    if [[ -n "$(port_listener "$port")" ]]; then
      echo "$name is listening on port $port."
      return 0
    fi
    sleep 1
  done

  echo "$name did not start on port $port in time. Recent tmux output:" >&2
  tmux capture-pane -p -t "$SESSION_NAME:$window" | tail -n 80 >&2 || true
  exit 1
}

if tmux has-session -t "$SESSION_NAME" 2>/dev/null; then
  echo "Restarting existing tmux session: $SESSION_NAME"
  tmux kill-session -t "$SESSION_NAME"
  sleep 1
fi

ensure_port_free "$BACKEND_PORT" "backend"
ensure_port_free "$FRONTEND_PORT" "frontend"

tmux new-session -d -s "$SESSION_NAME" -n backend -c "$BACKEND_DIR" \
  "exec \"$BACKEND_PYTHON\" app.py"

tmux new-window -t "$SESSION_NAME" -n frontend -c "$FRONTEND_DIR" \
  "exec npm run serve"

wait_for_port "$BACKEND_PORT" "Backend" "backend"
wait_for_port "$FRONTEND_PORT" "Frontend" "frontend"

echo
echo "Started handwriting-web in tmux session: $SESSION_NAME"
echo "Frontend: http://localhost:$FRONTEND_PORT/"
echo "Backend:  http://127.0.0.1:$BACKEND_PORT"
echo
echo "View logs: tmux attach -t $SESSION_NAME"
echo "Stop:      ./start-dev.sh --stop"

if [[ "${1:-}" == "--attach" ]]; then
  exec tmux attach -t "$SESSION_NAME"
fi
