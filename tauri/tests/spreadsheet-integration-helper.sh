#!/usr/bin/env bash
# Helper script to manage the spreadsheet-poc web server for integration tests
# Addresses TSR-like behavior where the process doesn't terminate cleanly

set -e

ACTION="$1"
PORT="${2:-8085}"
SPREADSHEET_DIR="/home/user/RexxJS/examples/spreadsheet-poc"
PID_FILE="/tmp/spreadsheet-server-${PORT}.pid"
LOG_FILE="/tmp/spreadsheet-server-${PORT}.log"

# Kill process tree reliably
kill_process_tree() {
    local pid=$1
    local children=$(pgrep -P $pid 2>/dev/null || true)

    # Kill children first (recursive)
    for child in $children; do
        kill_process_tree $child
    done

    # Kill the process itself
    if ps -p $pid > /dev/null 2>&1; then
        kill -TERM $pid 2>/dev/null || true
        sleep 0.5

        # Force kill if still alive
        if ps -p $pid > /dev/null 2>&1; then
            kill -KILL $pid 2>/dev/null || true
        fi
    fi
}

case "$ACTION" in
    start)
        # Check if already running
        if [[ -f "$PID_FILE" ]]; then
            OLD_PID=$(cat "$PID_FILE")
            if ps -p $OLD_PID > /dev/null 2>&1; then
                echo "Server already running on port $PORT (PID: $OLD_PID)"
                exit 0
            fi
            rm -f "$PID_FILE"
        fi

        # Check if port is in use
        if lsof -Pi :$PORT -sTCP:LISTEN -t >/dev/null 2>&1; then
            echo "Error: Port $PORT is already in use"
            lsof -Pi :$PORT -sTCP:LISTEN
            exit 1
        fi

        echo "Starting spreadsheet web server on port $PORT..."

        # Start server in background
        cd "$SPREADSHEET_DIR"
        nohup npx http-server -p $PORT -c-1 --silent > "$LOG_FILE" 2>&1 &
        SERVER_PID=$!

        # Save PID
        echo $SERVER_PID > "$PID_FILE"

        # Wait for server to be ready
        for i in {1..30}; do
            if curl -s "http://localhost:$PORT/" > /dev/null 2>&1; then
                echo "✓ Server started successfully (PID: $SERVER_PID)"
                echo "  URL: http://localhost:$PORT/"
                echo "  Log: $LOG_FILE"
                exit 0
            fi
            sleep 0.5
        done

        echo "Error: Server failed to start"
        cat "$LOG_FILE"
        exit 1
        ;;

    stop)
        if [[ ! -f "$PID_FILE" ]]; then
            echo "No PID file found. Checking for processes on port $PORT..."

            # Try to find and kill processes using the port
            PORT_PIDS=$(lsof -ti :$PORT 2>/dev/null || true)
            if [[ -n "$PORT_PIDS" ]]; then
                echo "Found processes on port $PORT: $PORT_PIDS"
                for pid in $PORT_PIDS; do
                    echo "Killing process tree for PID $pid..."
                    kill_process_tree $pid
                done
                echo "✓ Killed all processes on port $PORT"
            else
                echo "No server running on port $PORT"
            fi
            exit 0
        fi

        SERVER_PID=$(cat "$PID_FILE")
        echo "Stopping server (PID: $SERVER_PID)..."

        # Kill process tree
        kill_process_tree $SERVER_PID

        # Clean up PID file
        rm -f "$PID_FILE"

        # Verify port is free
        sleep 1
        if lsof -Pi :$PORT -sTCP:LISTEN -t >/dev/null 2>&1; then
            echo "Warning: Port $PORT still in use after cleanup!"
            lsof -Pi :$PORT -sTCP:LISTEN
            exit 1
        fi

        echo "✓ Server stopped and port $PORT is free"
        rm -f "$LOG_FILE"
        ;;

    status)
        if [[ -f "$PID_FILE" ]]; then
            SERVER_PID=$(cat "$PID_FILE")
            if ps -p $SERVER_PID > /dev/null 2>&1; then
                echo "Server running (PID: $SERVER_PID)"
                echo "URL: http://localhost:$PORT/"
                exit 0
            else
                echo "PID file exists but process is dead"
                rm -f "$PID_FILE"
                exit 1
            fi
        else
            echo "Server not running"
            exit 1
        fi
        ;;

    restart)
        $0 stop $PORT
        sleep 2
        $0 start $PORT
        ;;

    *)
        echo "Usage: $0 {start|stop|status|restart} [port]"
        echo ""
        echo "Manages the spreadsheet-poc web server for integration tests"
        echo ""
        echo "Commands:"
        echo "  start [port]   - Start server (default port: 8085)"
        echo "  stop [port]    - Stop server and clean up properly"
        echo "  status [port]  - Check if server is running"
        echo "  restart [port] - Restart server"
        echo ""
        echo "This script properly handles process cleanup to avoid TSR-like behavior"
        exit 1
        ;;
esac
