-- SQANDAL RADAR schema (sqlite). Сгенерировано из sqandal/models.py
CREATE TABLE app_state (
	"key" VARCHAR(64) NOT NULL, 
	value JSON NOT NULL, 
	updated_at DATETIME NOT NULL, 
	CONSTRAINT pk_app_state PRIMARY KEY ("key")
);


CREATE TABLE sources (
	id INTEGER NOT NULL, 
	"key" VARCHAR(64) NOT NULL, 
	name VARCHAR(128) NOT NULL, 
	kind VARCHAR(24) NOT NULL, 
	scene VARCHAR(24) NOT NULL, 
	url TEXT NOT NULL, 
	enabled BOOLEAN NOT NULL, 
	config JSON NOT NULL, 
	state JSON NOT NULL, 
	etag TEXT, 
	last_modified TEXT, 
	tg_last_message_id BIGINT, 
	last_fetch_at DATETIME, 
	last_success_at DATETIME, 
	last_error TEXT, 
	consecutive_failures INTEGER NOT NULL, 
	items_total INTEGER NOT NULL, 
	created_at DATETIME NOT NULL, 
	updated_at DATETIME NOT NULL, 
	CONSTRAINT pk_sources PRIMARY KEY (id), 
	CONSTRAINT uq_sources_key UNIQUE ("key"), 
	CONSTRAINT ck_sources_source_kind CHECK (kind IN ('rss', 'web', 'telegram')), 
	CONSTRAINT ck_sources_scene CHECK (scene IN ('ru', 'de', 'us', 'intl'))
);


CREATE TABLE clusters (
	id INTEGER NOT NULL, 
	title TEXT NOT NULL, 
	scene VARCHAR(24) NOT NULL, 
	scenes JSON NOT NULL, 
	tags JSON NOT NULL, 
	artists JSON NOT NULL, 
	priority VARCHAR(24) NOT NULL, 
	priority_score INTEGER NOT NULL, 
	status VARCHAR(24) NOT NULL, 
	origin_item_id BIGINT, 
	origin_source_id INTEGER, 
	origin_published_at DATETIME NOT NULL, 
	origin_date_confidence VARCHAR(24) NOT NULL, 
	first_seen_at DATETIME NOT NULL, 
	last_item_at DATETIME NOT NULL, 
	items_count INTEGER NOT NULL, 
	sources_count INTEGER NOT NULL, 
	verification VARCHAR(24) NOT NULL, 
	is_stale BOOLEAN NOT NULL, 
	stale_reason TEXT, 
	republish_of_cluster_id BIGINT, 
	needs_refresh BOOLEAN NOT NULL, 
	centroid BLOB, 
	centroid_model VARCHAR(160), 
	created_at DATETIME NOT NULL, 
	updated_at DATETIME NOT NULL, 
	CONSTRAINT pk_clusters PRIMARY KEY (id), 
	CONSTRAINT ck_clusters_scene CHECK (scene IN ('ru', 'de', 'us', 'intl')), 
	CONSTRAINT ck_clusters_priority CHECK (priority IN ('CRITICAL', 'HIGH', 'NORMAL')), 
	CONSTRAINT ck_clusters_cluster_status CHECK (status IN ('new', 'drafted', 'scheduled', 'published', 'dismissed')), 
	CONSTRAINT fk_clusters_origin_source_id_sources FOREIGN KEY(origin_source_id) REFERENCES sources (id) ON DELETE SET NULL, 
	CONSTRAINT ck_clusters_confidence CHECK (origin_date_confidence IN ('high', 'medium', 'low')), 
	CONSTRAINT ck_clusters_verification CHECK (verification IN ('single_source', 'multi_source', 'insider_only'))
);

CREATE INDEX ix_clusters_origin_published ON clusters (origin_published_at);
CREATE INDEX ix_clusters_scene_last_item ON clusters (scene, last_item_at);
CREATE INDEX ix_clusters_status_priority ON clusters (status, priority);

CREATE TABLE fetch_runs (
	id INTEGER NOT NULL, 
	source_id INTEGER NOT NULL, 
	started_at DATETIME NOT NULL, 
	finished_at DATETIME, 
	ok BOOLEAN NOT NULL, 
	items_seen INTEGER NOT NULL, 
	items_new INTEGER NOT NULL, 
	error TEXT, 
	duration_ms INTEGER, 
	CONSTRAINT pk_fetch_runs PRIMARY KEY (id), 
	CONSTRAINT fk_fetch_runs_source_id_sources FOREIGN KEY(source_id) REFERENCES sources (id) ON DELETE CASCADE
);

CREATE INDEX ix_fetch_runs_source_started ON fetch_runs (source_id, started_at);

CREATE TABLE news_items (
	id INTEGER NOT NULL, 
	source_id INTEGER NOT NULL, 
	external_id VARCHAR(512) NOT NULL, 
	url TEXT, 
	canonical_url TEXT, 
	title TEXT NOT NULL, 
	summary TEXT, 
	content TEXT, 
	author VARCHAR(256), 
	image_url TEXT, 
	lang VARCHAR(8), 
	scene VARCHAR(24) NOT NULL, 
	published_at DATETIME, 
	date_source VARCHAR(24) NOT NULL, 
	date_confidence VARCHAR(24) NOT NULL, 
	fetched_at DATETIME NOT NULL, 
	effective_at DATETIME NOT NULL, 
	content_hash VARCHAR(64) NOT NULL, 
	embedding BLOB, 
	embedding_model VARCHAR(160), 
	cluster_id INTEGER, 
	similarity FLOAT, 
	is_relevant BOOLEAN NOT NULL, 
	tags JSON NOT NULL, 
	tg_channel VARCHAR(128), 
	tg_message_id BIGINT, 
	tg_views INTEGER, 
	fwd_from VARCHAR(256), 
	extra JSON NOT NULL, 
	created_at DATETIME NOT NULL, 
	CONSTRAINT pk_news_items PRIMARY KEY (id), 
	CONSTRAINT uq_news_items_source_external UNIQUE (source_id, external_id), 
	CONSTRAINT fk_news_items_source_id_sources FOREIGN KEY(source_id) REFERENCES sources (id) ON DELETE CASCADE, 
	CONSTRAINT ck_news_items_scene CHECK (scene IN ('ru', 'de', 'us', 'intl')), 
	CONSTRAINT ck_news_items_date_source CHECK (date_source IN ('feed_published', 'feed_updated', 'jsonld', 'meta', 'time_tag', 'url', 'telegram', 'telegram_forward', 'page_day', 'fetched')), 
	CONSTRAINT ck_news_items_confidence CHECK (date_confidence IN ('high', 'medium', 'low')), 
	CONSTRAINT fk_news_items_cluster_id_clusters FOREIGN KEY(cluster_id) REFERENCES clusters (id) ON DELETE SET NULL
);

CREATE INDEX ix_news_items_canonical_url ON news_items (canonical_url);
CREATE INDEX ix_news_items_cluster ON news_items (cluster_id);
CREATE INDEX ix_news_items_content_hash ON news_items (content_hash);
CREATE INDEX ix_news_items_effective_at ON news_items (effective_at);

CREATE TABLE posts (
	id INTEGER NOT NULL, 
	cluster_id INTEGER NOT NULL, 
	version INTEGER NOT NULL, 
	mode VARCHAR(24) NOT NULL, 
	status VARCHAR(24) NOT NULL, 
	flag_emoji VARCHAR(32) NOT NULL, 
	topic_emoji VARCHAR(32) NOT NULL, 
	headline TEXT NOT NULL, 
	bullets JSON NOT NULL, 
	summary TEXT NOT NULL, 
	tags JSON NOT NULL, 
	sources JSON NOT NULL, 
	footnotes JSON NOT NULL, 
	body_html TEXT NOT NULL, 
	body_text TEXT NOT NULL, 
	warnings JSON NOT NULL, 
	editor_notes TEXT, 
	is_publishable BOOLEAN NOT NULL, 
	event_date VARCHAR(10), 
	llm_model VARCHAR(64), 
	input_tokens INTEGER, 
	output_tokens INTEGER, 
	cache_read_tokens INTEGER, 
	scheduled_at DATETIME, 
	published_at DATETIME, 
	tg_chat_id VARCHAR(64), 
	tg_message_id BIGINT, 
	error TEXT, 
	created_at DATETIME NOT NULL, 
	updated_at DATETIME NOT NULL, 
	CONSTRAINT pk_posts PRIMARY KEY (id), 
	CONSTRAINT uq_posts_cluster_version UNIQUE (cluster_id, version), 
	CONSTRAINT fk_posts_cluster_id_clusters FOREIGN KEY(cluster_id) REFERENCES clusters (id) ON DELETE CASCADE, 
	CONSTRAINT ck_posts_post_mode CHECK (mode IN ('standard', 'harder', 'hardest', 'manual')), 
	CONSTRAINT ck_posts_post_status CHECK (status IN ('draft', 'scheduled', 'publishing', 'published', 'failed', 'superseded'))
);

CREATE INDEX ix_posts_status_scheduled ON posts (status, scheduled_at);

