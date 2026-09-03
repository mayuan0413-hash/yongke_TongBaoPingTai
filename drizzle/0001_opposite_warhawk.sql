CREATE TABLE `data_sources` (
	`id` text PRIMARY KEY NOT NULL,
	`project_id` text NOT NULL,
	`name` text NOT NULL,
	`connection_key` text NOT NULL,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	FOREIGN KEY (`project_id`) REFERENCES `projects`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE INDEX `idx_data_sources_project` ON `data_sources` (`project_id`);--> statement-breakpoint
CREATE TABLE `sheet_data_blocks` (
	`sheet_id` text NOT NULL,
	`block_index` integer NOT NULL,
	`rows` text NOT NULL,
	PRIMARY KEY(`sheet_id`, `block_index`),
	FOREIGN KEY (`sheet_id`) REFERENCES `sheets`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE TABLE `sheet_data_queries` (
	`sheet_id` text PRIMARY KEY NOT NULL,
	`data_source_id` text NOT NULL,
	`table_name` text NOT NULL,
	`selected_fields` text NOT NULL,
	`filters` text DEFAULT '[]' NOT NULL,
	`order_by` text DEFAULT '[]' NOT NULL,
	`row_limit` integer DEFAULT 10000 NOT NULL,
	`columns` text DEFAULT '[]' NOT NULL,
	`last_refreshed_at` text,
	`last_row_count` integer DEFAULT 0 NOT NULL,
	`truncated` integer DEFAULT false NOT NULL,
	FOREIGN KEY (`sheet_id`) REFERENCES `sheets`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`data_source_id`) REFERENCES `data_sources`(`id`) ON UPDATE no action ON DELETE restrict
);
--> statement-breakpoint
CREATE INDEX `idx_sheet_queries_data_source` ON `sheet_data_queries` (`data_source_id`);