CREATE TABLE `projects` (
	`id` text PRIMARY KEY NOT NULL,
	`name` text NOT NULL,
	`revision` integer DEFAULT 0 NOT NULL,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL
);
--> statement-breakpoint
CREATE TABLE `sheet_rows` (
	`sheet_id` text NOT NULL,
	`row_index` integer NOT NULL,
	`cells` text NOT NULL,
	PRIMARY KEY(`sheet_id`, `row_index`),
	FOREIGN KEY (`sheet_id`) REFERENCES `sheets`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE TABLE `sheets` (
	`id` text PRIMARY KEY NOT NULL,
	`project_id` text NOT NULL,
	`name` text NOT NULL,
	`position` integer NOT NULL,
	`row_count` integer NOT NULL,
	`column_count` integer NOT NULL,
	`row_heights` text DEFAULT '{}' NOT NULL,
	`column_widths` text DEFAULT '{}' NOT NULL,
	FOREIGN KEY (`project_id`) REFERENCES `projects`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE INDEX `idx_sheets_project_position` ON `sheets` (`project_id`,`position`);