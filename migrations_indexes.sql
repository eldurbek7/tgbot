-- Railway/PostgreSQL performance indexes for Telegram bot callbacks.
-- Safe to run multiple times.

CREATE INDEX IF NOT EXISTS idx_users_user_id ON user_prefs(user_id);
CREATE INDEX IF NOT EXISTS idx_votes_user_id ON votes(user_id);
CREATE INDEX IF NOT EXISTS idx_votes_subject_key ON votes(subject_key);
CREATE INDEX IF NOT EXISTS idx_votes_teacher_key ON votes(teacher_key);
CREATE INDEX IF NOT EXISTS idx_votes_subject_teacher ON votes(subject_key, teacher_key);
CREATE INDEX IF NOT EXISTS idx_votes_voted_at ON votes(voted_at DESC);
CREATE INDEX IF NOT EXISTS idx_teacher_ratings_user_id ON teacher_ratings(user_id);
CREATE INDEX IF NOT EXISTS idx_teacher_ratings_subject_teacher ON teacher_ratings(subject_key, teacher_key);
CREATE INDEX IF NOT EXISTS idx_teacher_ratings_rated_at ON teacher_ratings(rated_at DESC);
CREATE INDEX IF NOT EXISTS idx_complaints_user_id ON complaints(user_id);
CREATE INDEX IF NOT EXISTS idx_complaints_type_id ON complaints(type, id DESC);
CREATE INDEX IF NOT EXISTS idx_complaints_created_at ON complaints(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_db_subjects_sort_order ON db_subjects(sort_order, subject_key);
CREATE INDEX IF NOT EXISTS idx_db_teachers_subject_key ON db_teachers(subject_key);
CREATE INDEX IF NOT EXISTS idx_db_teachers_teacher_key ON db_teachers(teacher_key);
