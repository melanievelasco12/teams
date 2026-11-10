<?php

declare(strict_types=1);

namespace TeamsBrightspace;

use PDO;

final class CourseSessionRepository
{
    private PDO $database;

    public function __construct(?string $databasePath = null)
    {
        $databasePath ??= dirname(__DIR__).'/storage/course-sessions.sqlite';
        $directory = dirname($databasePath);
        if (!is_dir($directory) && !mkdir($directory, 0700, true) && !is_dir($directory)) {
            throw new \RuntimeException('No se pudo preparar el almacenamiento de sesiones.');
        }

        $this->database = new PDO('sqlite:'.$databasePath, null, null, [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        ]);
        $this->database->exec('PRAGMA foreign_keys = ON');
        $this->database->exec('PRAGMA journal_mode = WAL');
        $this->database->exec(
            'CREATE TABLE IF NOT EXISTS course_sessions (
                id TEXT PRIMARY KEY,
                context_id TEXT NOT NULL,
                graph_event_id TEXT NOT NULL,
                title TEXT NOT NULL,
                start_at TEXT NOT NULL,
                end_at TEXT NOT NULL,
                time_zone TEXT NOT NULL,
                join_url TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE (context_id, graph_event_id)
            )'
        );
        $this->database->exec(
            'CREATE INDEX IF NOT EXISTS course_sessions_context_start
             ON course_sessions (context_id, start_at)'
        );
        $this->normalizeLegacyTimes();
    }

    public function publish(string $contextId, array $session): array
    {
        $existing = $this->database->prepare(
            'SELECT id, graph_event_id, title, start_at, end_at, time_zone, join_url
             FROM course_sessions WHERE context_id = :context_id AND graph_event_id = :graph_event_id'
        );
        $existing->execute([
            'context_id' => $contextId,
            'graph_event_id' => $session['graphEventId'],
        ]);
        $record = $existing->fetch();
        if (is_array($record)) {
            if ($record['title'] !== $session['title']
                || $record['start_at'] !== $session['start']
                || $record['end_at'] !== $session['end']
                || $record['time_zone'] !== $session['timeZone']
                || $record['join_url'] !== $session['joinUrl']) {
                throw new HttpException(409, 'Esta reunión ya está publicada con otros datos.');
            }

            return $this->format($record);
        }

        $id = bin2hex(random_bytes(16));
        $insert = $this->database->prepare(
            'INSERT OR IGNORE INTO course_sessions
             (id, context_id, graph_event_id, title, start_at, end_at, time_zone, join_url, created_at)
             VALUES (:id, :context_id, :graph_event_id, :title, :start_at, :end_at, :time_zone, :join_url, :created_at)'
        );
        $insert->execute([
            'id' => $id,
            'context_id' => $contextId,
            'graph_event_id' => $session['graphEventId'],
            'title' => $session['title'],
            'start_at' => $session['start'],
            'end_at' => $session['end'],
            'time_zone' => $session['timeZone'],
            'join_url' => $session['joinUrl'],
            'created_at' => gmdate(DATE_ATOM),
        ]);
        if ($insert->rowCount() === 0) {
            $existing->execute([
                'context_id' => $contextId,
                'graph_event_id' => $session['graphEventId'],
            ]);
            $record = $existing->fetch();
            if (!is_array($record)
                || $record['title'] !== $session['title']
                || $record['start_at'] !== $session['start']
                || $record['end_at'] !== $session['end']
                || $record['time_zone'] !== $session['timeZone']
                || $record['join_url'] !== $session['joinUrl']) {
                throw new HttpException(409, 'Esta reunión ya está publicada con otros datos.');
            }

            return $this->format($record);
        }

        return [
            'id' => $id,
            'graphEventId' => $session['graphEventId'],
            'start' => $session['start'],
            'end' => $session['end'],
            'timeZone' => $session['timeZone'],
            'topic' => $session['title'],
            'joinUrl' => $session['joinUrl'],
        ];
    }

    public function listForCourse(string $contextId): array
    {
        $query = $this->database->prepare(
            'SELECT id, graph_event_id, start_at, end_at, time_zone, title, join_url
             FROM course_sessions WHERE context_id = :context_id ORDER BY start_at'
        );
        $query->execute(['context_id' => $contextId]);

        return array_map(fn (array $record): array => $this->format($record), $query->fetchAll());
    }

    public function deleteForCourse(string $contextId, string $sessionId): bool
    {
        $delete = $this->database->prepare(
            'DELETE FROM course_sessions WHERE id = :id AND context_id = :context_id'
        );
        $delete->execute([
            'id' => $sessionId,
            'context_id' => $contextId,
        ]);

        return $delete->rowCount() > 0;
    }

    private function format(array $record): array
    {
        return [
            'id' => $record['id'],
            'graphEventId' => $record['graph_event_id'],
            'start' => $record['start_at'],
            'end' => $record['end_at'],
            'timeZone' => $record['time_zone'],
            'topic' => $record['title'],
            'joinUrl' => $record['join_url'],
        ];
    }

    private function normalizeLegacyTimes(): void
    {
        $this->database->exec(
            'CREATE TABLE IF NOT EXISTS course_session_meta (
                meta_key TEXT PRIMARY KEY,
                meta_value TEXT NOT NULL
            )'
        );
        $check = $this->database->prepare(
            'SELECT meta_value FROM course_session_meta WHERE meta_key = :meta_key'
        );
        $check->execute(['meta_key' => 'timezone_normalization_v1']);
        if ($check->fetchColumn() !== false) {
            return;
        }

        $windowsTimeZones = [
            'Central Standard Time (Mexico)' => 'America/Mexico_City',
            'Pacific Standard Time' => 'America/Los_Angeles',
            'Mountain Standard Time' => 'America/Denver',
            'Central Standard Time' => 'America/Chicago',
            'Eastern Standard Time' => 'America/New_York',
            'UTC' => 'UTC',
        ];
        $legacyZone = new \DateTimeZone(date_default_timezone_get());
        $this->database->beginTransaction();
        try {
            $records = $this->database->query(
                'SELECT id, start_at, end_at, time_zone FROM course_sessions'
            )->fetchAll();
            $update = $this->database->prepare(
                'UPDATE course_sessions SET start_at = :start_at, end_at = :end_at WHERE id = :id'
            );
            foreach ($records as $record) {
                $sessionZone = new \DateTimeZone(
                    $windowsTimeZones[$record['time_zone']] ?? $record['time_zone']
                );
                $startWallTime = (new \DateTimeImmutable($record['start_at']))
                    ->setTimezone($legacyZone)
                    ->format('Y-m-d\TH:i:s');
                $endWallTime = (new \DateTimeImmutable($record['end_at']))
                    ->setTimezone($legacyZone)
                    ->format('Y-m-d\TH:i:s');
                $start = new \DateTimeImmutable($startWallTime, $sessionZone);
                $end = new \DateTimeImmutable($endWallTime, $sessionZone);
                $update->execute([
                    'id' => $record['id'],
                    'start_at' => gmdate(DATE_ATOM, $start->getTimestamp()),
                    'end_at' => gmdate(DATE_ATOM, $end->getTimestamp()),
                ]);
            }
            $mark = $this->database->prepare(
                'INSERT INTO course_session_meta (meta_key, meta_value)
                 VALUES (:meta_key, :meta_value)'
            );
            $mark->execute([
                'meta_key' => 'timezone_normalization_v1',
                'meta_value' => gmdate(DATE_ATOM),
            ]);
            $this->database->commit();
        } catch (\Throwable $error) {
            $this->database->rollBack();
            throw $error;
        }
    }
}
