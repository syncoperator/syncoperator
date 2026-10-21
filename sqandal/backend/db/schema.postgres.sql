-- SQANDAL RADAR schema (postgresql). Сгенерировано из sqandal/models.py
CREATE TABLE app_state (
	key VARCHAR(64) NOT NULL, 
	value JSONB NOT NULL, 
	updated_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	CONSTRAINT pk_app_state PRIMARY KEY (key)
);

COMMENT ON TABLE app_state IS 'Служебные ключи: команды воркеру, метки';

CREATE TABLE sources (
	id BIGSERIAL NOT NULL, 
	key VARCHAR(64) NOT NULL, 
	name VARCHAR(128) NOT NULL, 
	kind VARCHAR(24) NOT NULL, 
	scene VARCHAR(24) NOT NULL, 
	url TEXT NOT NULL, 
	enabled BOOLEAN NOT NULL, 
	config JSONB NOT NULL, 
	state JSONB NOT NULL, 
	etag TEXT, 
	last_modified TEXT, 
	tg_last_message_id BIGINT, 
	last_fetch_at TIMESTAMP WITH TIME ZONE, 
	last_success_at TIMESTAMP WITH TIME ZONE, 
	last_error TEXT, 
	consecutive_failures INTEGER NOT NULL, 
	items_total INTEGER NOT NULL, 
	created_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	updated_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	CONSTRAINT pk_sources PRIMARY KEY (id), 
	CONSTRAINT uq_sources_key UNIQUE (key), 
	CONSTRAINT ck_sources_source_kind CHECK (kind IN ('rss', 'web', 'telegram')), 
	CONSTRAINT ck_sources_scene CHECK (scene IN ('ru', 'de', 'us', 'intl'))
);

COMMENT ON TABLE sources IS 'Источники новостей: RSS, сайты, Telegram-каналы';
COMMENT ON COLUMN sources.key IS 'Слаг из sources.yaml';
COMMENT ON COLUMN sources.state IS 'Состояние сборщика (найденный фид, курсоры)';
COMMENT ON COLUMN sources.tg_last_message_id IS 'Курсор догона Telegram';

CREATE TABLE clusters (
	id BIGSERIAL NOT NULL, 
	title TEXT NOT NULL, 
	scene VARCHAR(24) NOT NULL, 
	scenes JSONB NOT NULL, 
	tags JSONB NOT NULL, 
	artists JSONB NOT NULL, 
	priority VARCHAR(24) NOT NULL, 
	priority_score INTEGER NOT NULL, 
	status VARCHAR(24) NOT NULL, 
	origin_item_id BIGINT, 
	origin_source_id BIGINT, 
	origin_published_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	origin_date_confidence VARCHAR(24) NOT NULL, 
	first_seen_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	last_item_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	items_count INTEGER NOT NULL, 
	sources_count INTEGER NOT NULL, 
	verification VARCHAR(24) NOT NULL, 
	is_stale BOOLEAN NOT NULL, 
	stale_reason TEXT, 
	republish_of_cluster_id BIGINT, 
	needs_refresh BOOLEAN NOT NULL, 
	centroid BYTEA, 
	centroid_model VARCHAR(160), 
	created_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	updated_at TIMESTAMP WITH TIME ZONE NOT NULL, 
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
COMMENT ON TABLE clusters IS 'Сюжеты: кластеры новостей об одном событии';
COMMENT ON COLUMN clusters.origin_item_id IS 'Самая ранняя публикация сюжета';
COMMENT ON COLUMN clusters.origin_published_at IS 'Дата/время первоисточника (min по сюжету)';
COMMENT ON COLUMN clusters.republish_of_cluster_id IS 'Похож на уже опубликованный сюжет';
COMMENT ON COLUMN clusters.needs_refresh IS 'Появились новые источники после черновика';

CREATE TABLE fetch_runs (
	id BIGSERIAL NOT NULL, 
	source_id BIGINT NOT NULL, 
	started_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	finished_at TIMESTAMP WITH TIME ZONE, 
	ok BOOLEAN NOT NULL, 
	items_seen INTEGER NOT NULL, 
	items_new INTEGER NOT NULL, 
	error TEXT, 
	duration_ms INTEGER, 
	CONSTRAINT pk_fetch_runs PRIMARY KEY (id), 
	CONSTRAINT fk_fetch_runs_source_id_sources FOREIGN KEY(source_id) REFERENCES sources (id) ON DELETE CASCADE
);

CREATE INDEX ix_fetch_runs_source_started ON fetch_runs (source_id, started_at);
COMMENT ON TABLE fetch_runs IS 'Журнал прогонов сборщиков';

CREATE TABLE news_items (
	id BIGSERIAL NOT NULL, 
	source_id BIGINT NOT NULL, 
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
	published_at TIMESTAMP WITH TIME ZONE, 
	date_source VARCHAR(24) NOT NULL, 
	date_confidence VARCHAR(24) NOT NULL, 
	fetched_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	effective_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	content_hash VARCHAR(64) NOT NULL, 
	embedding BYTEA, 
	embedding_model VARCHAR(160), 
	cluster_id BIGINT, 
	similarity FLOAT, 
	is_relevant BOOLEAN NOT NULL, 
	tags JSONB NOT NULL, 
	tg_channel VARCHAR(128), 
	tg_message_id BIGINT, 
	tg_views INTEGER, 
	fwd_from VARCHAR(256), 
	extra JSONB NOT NULL, 
	created_at TIMESTAMP WITH TIME ZONE NOT NULL, 
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
COMMENT ON TABLE news_items IS 'Сырые публикации из источников';
COMMENT ON COLUMN news_items.external_id IS 'guid / url / channel:msg_id';
COMMENT ON COLUMN news_items.content IS 'Полный текст (если скачан)';
COMMENT ON COLUMN news_items.published_at IS 'Дата, заявленная источником';
COMMENT ON COLUMN news_items.effective_at IS 'Проверенное время публикации (для первоисточника)';
COMMENT ON COLUMN news_items.fwd_from IS 'Исходный канал/автор репоста';

CREATE TABLE posts (
	id BIGSERIAL NOT NULL, 
	cluster_id BIGINT NOT NULL, 
	version INTEGER NOT NULL, 
	mode VARCHAR(24) NOT NULL, 
	status VARCHAR(24) NOT NULL, 
	flag_emoji VARCHAR(32) NOT NULL, 
	topic_emoji VARCHAR(32) NOT NULL, 
	headline TEXT NOT NULL, 
	bullets JSONB NOT NULL, 
	summary TEXT NOT NULL, 
	tags JSONB NOT NULL, 
	sources JSONB NOT NULL, 
	footnotes JSONB NOT NULL, 
	body_html TEXT NOT NULL, 
	body_text TEXT NOT NULL, 
	warnings JSONB NOT NULL, 
	editor_notes TEXT, 
	is_publishable BOOLEAN NOT NULL, 
	event_date VARCHAR(10), 
	llm_model VARCHAR(64), 
	input_tokens INTEGER, 
	output_tokens INTEGER, 
	cache_read_tokens INTEGER, 
	scheduled_at TIMESTAMP WITH TIME ZONE, 
	published_at TIMESTAMP WITH TIME ZONE, 
	tg_chat_id VARCHAR(64), 
	tg_message_id BIGINT, 
	error TEXT, 
	created_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	updated_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	CONSTRAINT pk_posts PRIMARY KEY (id), 
	CONSTRAINT uq_posts_cluster_version UNIQUE (cluster_id, version), 
	CONSTRAINT fk_posts_cluster_id_clusters FOREIGN KEY(cluster_id) REFERENCES clusters (id) ON DELETE CASCADE, 
	CONSTRAINT ck_posts_post_mode CHECK (mode IN ('standard', 'harder', 'hardest', 'manual')), 
	CONSTRAINT ck_posts_post_status CHECK (status IN ('draft', 'scheduled', 'publishing', 'published', 'failed', 'superseded'))
);

CREATE INDEX ix_posts_status_scheduled ON posts (status, scheduled_at);
COMMENT ON TABLE posts IS 'ТГ-посты: версии генерации, правки, публикации';
COMMENT ON COLUMN posts.sources IS '[{name, url}] для строки источников';

