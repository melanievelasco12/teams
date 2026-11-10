<?php

declare(strict_types=1);

namespace TeamsBrightspace;

use PDO;

final class CourseRecordingRepository
{
    private PDO $database;

    public function __construct(?string $databasePath = null)
    {
        $databasePath ??= dirname(__DIR__).'/storage/course-sessions.sqlite';
        $this->database = new PDO('sqlite:'.$databasePath, null, null, [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        ]);
        $this->database->exec('PRAGMA foreign_keys = ON');
        $this->database->exec(
            'CREATE TABLE IF NOT EXISTS course_recordings (
                id TEXT PRIMARY KEY,
                context_id TEXT NOT NULL,
                course_session_id TEXT NOT NULL,
                drive_item_id TEXT NOT NULL,
                title TEXT NOT NULL,
                created_at TEXT NOT NULL,
                web_url TEXT NOT NULL,
                published_at TEXT NOT NULL,
                UNIQUE (context_id, drive_item_id),
                FOREIGN KEY (course_session_id) REFERENCES course_sessions (id) ON DELETE CASCADE
            )'
        );
        $this->database->exec(
            'CREATE INDEX IF NOT EXISTS course_recordings_context_created
             ON course_recordings (context_id, created_at)'
        );
    }

    public function publish(string $contextId, array $recording): array
    {
        $sessionQuery = $this->database->prepare(
            'SELECT id FROM course_sessions WHERE id = :session_id AND context_id = :context_id'
        );
        $sessionQuery->execute([
            'session_id' => $recording['sessionId'],
            'context_id' => $contextId,
        ]);
        if ($sessionQuery->fetchColumn() === false) {
            throw new HttpException(404, 'No se encontró la sesión de este curso para vincular la grabación.');
        }

        $existingQuery = $this->database->prepare(
            'SELECT id, course_session_id, drive_item_id, title, created_at, web_url
             FROM course_recordings WHERE context_id = :context_id AND drive_item_id = :drive_item_id'
        );
        $existingQuery->execute([
            'context_id' => $contextId,
            'drive_item_id' => $recording['driveItemId'],
        ]);
        $existing = $existingQuery->fetch();
        if (is_array($existing)) {
            if ($existing['course_session_id'] !== $recording['sessionId']) {
                throw new HttpException(409, 'Esta grabación ya está vinculada a otra sesión del curso.');
            }

            return $this->format($existing, true);
        }

        $id = bin2hex(random_bytes(16));
        $insert = $this->database->prepare(
            'INSERT OR IGNORE INTO course_recordings
             (id, context_id, course_session_id, drive_item_id, title, created_at, web_url, published_at)
             VALUES (:id, :context_id, :course_session_id, :drive_item_id, :title, :created_at, :web_url, :published_at)'
        );
        $insert->execute([
            'id' => $id,
            'context_id' => $contextId,
            'course_session_id' => $recording['sessionId'],
            'drive_item_id' => $recording['driveItemId'],
            'title' => $recording['title'],
            'created_at' => $recording['createdAt'],
            'web_url' => $recording['webUrl'],
            'published_at' => gmdate(DATE_ATOM),
        ]);
        if ($insert->rowCount() === 0) {
            $existingQuery->execute([
                'context_id' => $contextId,
                'drive_item_id' => $recording['driveItemId'],
            ]);
            $existing = $existingQuery->fetch();
            if (!is_array($existing) || $existing['course_session_id'] !== $recording['sessionId']) {
                throw new HttpException(409, 'Esta grabación ya está vinculada a otra sesión del curso.');
            }

            return $this->format($existing, true);
        }

        return [
            'id' => $id,
            'sessionId' => $recording['sessionId'],
            'driveItemId' => $recording['driveItemId'],
            'title' => $recording['title'],
            'createdAt' => $recording['createdAt'],
            'webUrl' => $recording['webUrl'],
        ];
    }

    public function listForCourse(string $contextId, bool $includeDriveItemIds = false): array
    {
        $query = $this->database->prepare(
            'SELECT id, course_session_id, drive_item_id, title, created_at, web_url
             FROM course_recordings WHERE context_id = :context_id ORDER BY created_at DESC'
        );
        $query->execute(['context_id' => $contextId]);

        return array_map(
            fn (array $record): array => $this->format($record, $includeDriveItemIds),
            $query->fetchAll()
        );
    }

    private function format(array $record, bool $includeDriveItemId): array
    {
        $formatted = [
            'id' => $record['id'],
            'sessionId' => $record['course_session_id'],
            'title' => $record['title'],
            'createdAt' => $record['created_at'],
            'webUrl' => $record['web_url'],
        ];
        if ($includeDriveItemId) {
            $formatted['driveItemId'] = $record['drive_item_id'];
        }

        return $formatted;
    }
}
