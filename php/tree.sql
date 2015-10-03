/*
Navicat MySQL Data Transfer

Source Server         : Prevent
Source Server Version : 50524
Source Host           : localhost:3306
Source Database       : test_tree

Target Server Type    : MYSQL
Target Server Version : 50524
File Encoding         : 65001

Date: 2013-07-03 16:39:57
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
-- Table structure for `tree`
-- ----------------------------
DROP TABLE IF EXISTS `tree`;
CREATE TABLE `tree` (
  `id` bigint(20) unsigned NOT NULL AUTO_INCREMENT,
  `parent_id` bigint(20) unsigned DEFAULT NULL,
  `foreign_id` bigint(20) unsigned DEFAULT NULL,
  `foreign_table` enum('users','images','posts') DEFAULT NULL,
  `child_table` enum('users','images','posts') DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `parent_id` (`parent_id`),
  KEY `foreign_id` (`foreign_id`,`foreign_table`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;

