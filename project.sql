-- phpMyAdmin SQL Dump
-- version 4.0.4.1
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: Feb 05, 2016 at 12:28 PM
-- Server version: 5.5.32
-- PHP Version: 5.4.19

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `project`
--
CREATE DATABASE IF NOT EXISTS `project` DEFAULT CHARACTER SET latin1 COLLATE latin1_swedish_ci;
USE `project`;

-- --------------------------------------------------------

--
-- Table structure for table `appoinment`
--

CREATE TABLE IF NOT EXISTS `appoinment` (
  `doctor_id` int(11) DEFAULT NULL,
  `patient_id` int(100) DEFAULT NULL,
  `appoinment_date` varchar(30) DEFAULT NULL,
  `description` varchar(50) DEFAULT NULL,
  KEY `doctor_id` (`doctor_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `appoinment`
--

INSERT INTO `appoinment` (`doctor_id`, `patient_id`, `appoinment_date`, `description`) VALUES
(22, 102, '2015-7-7', 'skin problem'),
(100, 12025, '2089-56-6', 'emo'),
(1251, 4582, '6-5-2', 'omo');

-- --------------------------------------------------------

--
-- Table structure for table `billing`
--

CREATE TABLE IF NOT EXISTS `billing` (
  `patient_id` int(11) DEFAULT NULL,
  `opd_bill_no` int(50) DEFAULT NULL,
  `date` varchar(30) DEFAULT NULL,
  `consulting_fee` int(100) DEFAULT NULL,
  `patient_name` varchar(30) DEFAULT NULL,
  `patient_detail` varchar(30) DEFAULT NULL,
  `total_amount` int(100) DEFAULT NULL,
  `concession_amount` int(100) DEFAULT NULL,
  `net_amount` int(100) DEFAULT NULL,
  `amount_paid` int(100) DEFAULT NULL,
  KEY `patient_id` (`patient_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `billing`
--

INSERT INTO `billing` (`patient_id`, `opd_bill_no`, `date`, `consulting_fee`, `patient_name`, `patient_detail`, `total_amount`, `concession_amount`, `net_amount`, `amount_paid`) VALUES
(1, 145, '2055-5-3', 0, 'susan', 'no', 5000, 400, 4600, 4600),
(11, 123, '2015/8/7', 0, 'susan', 'no', 8000, 500, 7500, 7500),
(1212, 123, '2015-8-7', 0, 'ram', 'no', 5000, 500, 4500, 4500);

-- --------------------------------------------------------

--
-- Table structure for table `blood_test`
--

CREATE TABLE IF NOT EXISTS `blood_test` (
  `registration_id` int(100) NOT NULL,
  `patient_name` varchar(20) DEFAULT NULL,
  `test_date` varchar(20) DEFAULT NULL,
  `age` int(20) DEFAULT NULL,
  `gender` varchar(10) DEFAULT NULL,
  `haemoglobin` int(50) DEFAULT NULL,
  `t_l_c` int(50) DEFAULT NULL,
  `neutrophills` int(50) DEFAULT NULL,
  `lymphacytes` int(50) DEFAULT NULL,
  `eoslnophil` int(50) DEFAULT NULL,
  `monocytes` int(50) DEFAULT NULL,
  `creatinine` int(50) DEFAULT NULL,
  `bosophils` int(50) DEFAULT NULL,
  `platelets` int(50) DEFAULT NULL,
  PRIMARY KEY (`registration_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `blood_test`
--

INSERT INTO `blood_test` (`registration_id`, `patient_name`, `test_date`, `age`, `gender`, `haemoglobin`, `t_l_c`, `neutrophills`, `lymphacytes`, `eoslnophil`, `monocytes`, `creatinine`, `bosophils`, `platelets`) VALUES
(12, 'susan', '2015-5-12', 19, 'male', 14, 5000, 50, 30, 3, 2, 60, 1, 1540),
(120, 'susan', '2015-5-12', 19, 'male', 14, 5000, 50, 30, 3, 2, 60, 1, 1540);

-- --------------------------------------------------------

--
-- Table structure for table `discharge`
--

CREATE TABLE IF NOT EXISTS `discharge` (
  `patient_id` int(11) DEFAULT NULL,
  `name` varchar(30) DEFAULT NULL,
  `department` varchar(30) DEFAULT NULL,
  `problem` varchar(30) DEFAULT NULL,
  `appointment` varchar(30) DEFAULT NULL,
  `time_admitted` varchar(30) DEFAULT NULL,
  `bill_no` int(50) DEFAULT NULL,
  `ward_no` int(100) DEFAULT NULL,
  `paid` varchar(50) DEFAULT NULL,
  KEY `patient_id` (`patient_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `discharge`
--

INSERT INTO `discharge` (`patient_id`, `name`, `department`, `problem`, `appointment`, `time_admitted`, `bill_no`, `ward_no`, `paid`) VALUES
(1, 'susan', 'homo', 'skin', 'no', '2056-8-9', 123, 102, '5000'),
(11, 'susan', 'no', 'skin', 'no', '2015/8/7', 123, 23, '7500');

-- --------------------------------------------------------

--
-- Table structure for table `doctor_account`
--

CREATE TABLE IF NOT EXISTS `doctor_account` (
  `doctor_id` int(11) DEFAULT NULL,
  `doctor_name` varchar(30) DEFAULT NULL,
  `amount` int(100) DEFAULT NULL,
  `date` varchar(30) DEFAULT NULL,
  `salary` int(100) DEFAULT NULL,
  `bonous` int(100) DEFAULT NULL,
  KEY `doctor_id` (`doctor_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `doctor_account`
--

INSERT INTO `doctor_account` (`doctor_id`, `doctor_name`, `amount`, `date`, `salary`, `bonous`) VALUES
(120, 'susan', 50000, '06-07-2015', 0, 50000),
(10, 'ram', 5000, '08-07-2015', 0, 5000);

-- --------------------------------------------------------

--
-- Table structure for table `doctor_info`
--

CREATE TABLE IF NOT EXISTS `doctor_info` (
  `doctor_id` int(100) NOT NULL,
  `doctor_name` varchar(30) DEFAULT NULL,
  `address` varchar(30) DEFAULT NULL,
  `gender` varchar(30) DEFAULT NULL,
  `ph_no` int(30) DEFAULT NULL,
  `department` varchar(30) DEFAULT NULL,
  `entry_date` varchar(30) DEFAULT NULL,
  `education_qualification` varchar(30) DEFAULT NULL,
  `image` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`doctor_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `doctor_info`
--

INSERT INTO `doctor_info` (`doctor_id`, `doctor_name`, `address`, `gender`, `ph_no`, `department`, `entry_date`, `education_qualification`, `image`) VALUES
(10, 'ram', 'ktm', 'male', 98, 'skin', '08-07-2015', 'MBBS', NULL),
(22, 'ram hari pandya', 'kalanki', 'male', 9845, 'bone', '07-07-2015', 'MBBS', NULL),
(100, 'ram', 'ktm', 'male', 98, 'skin', '08-07-2015', 'MBBS', NULL),
(120, 'susan', 'lamjung', 'male', 2147483647, 'skin', '06-07-2015', 'MBBS', NULL),
(1022, 'susan', 'lamjung', 'male', 2147483647, 'skin', '12-08-2015', 'MBBS', 'C:UsersHPDocumentssusansUsAnIm'),
(1251, 'ram', 'lamjung', 'male', 2147483647, 'head', '13-08-2015', 'MBBS', 'D:5.jpg');

-- --------------------------------------------------------

--
-- Table structure for table `doctor_list`
--

CREATE TABLE IF NOT EXISTS `doctor_list` (
  `Doctor_name` varchar(50) DEFAULT NULL,
  `doctor_field` varchar(50) DEFAULT NULL,
  `experience` int(20) DEFAULT NULL,
  `contact_no` int(20) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `doctor_list`
--

INSERT INTO `doctor_list` (`Doctor_name`, `doctor_field`, `experience`, `contact_no`) VALUES
('susan', 'skin', 2, 9846),
('sushan', 'head', 3, 98463695);

-- --------------------------------------------------------

--
-- Table structure for table `login`
--

CREATE TABLE IF NOT EXISTS `login` (
  `user_name` varchar(30) DEFAULT NULL,
  `password` varchar(30) NOT NULL,
  `account_type` varchar(30) DEFAULT NULL,
  `ph_no` varchar(30) DEFAULT NULL,
  `email` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`password`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `login`
--

INSERT INTO `login` (`user_name`, `password`, `account_type`, `ph_no`, `email`) VALUES
('susan', 'bhujel', 'admin', '88', 'adsf'),
('My Hospital', 'Hospital3', 'admin', '9846369513', 'bhujel.susan@yahoo.com'),
('My Hospital', 'Hospital4', 'user', '', ''),
('My Hospital', 'Hospital5', 'recption', '', '');

-- --------------------------------------------------------

--
-- Table structure for table `login1`
--

CREATE TABLE IF NOT EXISTS `login1` (
  `user_name` varchar(30) DEFAULT NULL,
  `password` varchar(30) NOT NULL,
  `account_type` varchar(30) DEFAULT NULL,
  `ph_no` varchar(30) DEFAULT NULL,
  `email` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`password`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `login1`
--

INSERT INTO `login1` (`user_name`, `password`, `account_type`, `ph_no`, `email`) VALUES
('My Hospital', 'Hospital3', 'admin', '', ''),
('My Hospital', 'Hospital4', 'user', '', ''),
('My Hospital', 'Hospital5', 'recption', '', '');

-- --------------------------------------------------------

--
-- Table structure for table `patient_entry`
--

CREATE TABLE IF NOT EXISTS `patient_entry` (
  `patient_id` int(10) NOT NULL,
  `patient_name` varchar(30) DEFAULT NULL,
  `patient_address` varchar(30) DEFAULT NULL,
  `martial_status` varchar(30) DEFAULT NULL,
  `religion` varchar(30) DEFAULT NULL,
  `father_or_husband_name` varchar(30) DEFAULT NULL,
  `registration_date` int(15) DEFAULT NULL,
  `city` varchar(30) DEFAULT NULL,
  `mb_number` varchar(30) DEFAULT NULL,
  `Gender` varchar(30) DEFAULT NULL,
  `age` int(10) DEFAULT NULL,
  `dr_name` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`patient_id`),
  UNIQUE KEY `reg_entry` (`patient_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `patient_entry`
--

INSERT INTO `patient_entry` (`patient_id`, `patient_name`, `patient_address`, `martial_status`, `religion`, `father_or_husband_name`, `registration_date`, `city`, `mb_number`, `Gender`, `age`, `dr_name`) VALUES
(1, 'susan', 'besishahar', 'unmarrid', 'hindu', 'ram', 2071, 'besishahar', '9846369513', 'male', 19, 'ramhari baral'),
(11, 'susan', 'besishahar', 'unmarrid', 'hindu', 'ram', 2071, 'besishahar', '9846369513', 'male', 19, 'ramhari baral'),
(1212, 'susan', 'lamjung', 'unmarrid', 'hindu', 'nar bahadur bhujel', 2056, 'besishahar', '9846099897', 'male', 12, 'ram hari parshai'),
(1258, 'ram', 'ktm', 'unma', 'hindu', 'no', 2056, 'ktm', '9844', 'male', 15, 'hari');

-- --------------------------------------------------------

--
-- Table structure for table `staff_account`
--

CREATE TABLE IF NOT EXISTS `staff_account` (
  `staff_id` int(11) DEFAULT NULL,
  `staff_name` varchar(50) DEFAULT NULL,
  `amount` int(100) DEFAULT NULL,
  `date` varchar(20) DEFAULT NULL,
  `salary` int(100) DEFAULT NULL,
  `bonous` int(100) DEFAULT NULL,
  KEY `staff_id` (`staff_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `staff_account`
--

INSERT INTO `staff_account` (`staff_id`, `staff_name`, `amount`, `date`, `salary`, `bonous`) VALUES
(22, 'ram', 4500, '16-08-2015', 4500, 0),
(222, 'susan', 5000, '16-08-2015', 0, 5000);

-- --------------------------------------------------------

--
-- Table structure for table `staff_info`
--

CREATE TABLE IF NOT EXISTS `staff_info` (
  `staff_id` int(10) NOT NULL,
  `first_name` varchar(30) DEFAULT NULL,
  `last_name` varchar(30) DEFAULT NULL,
  `address` varchar(30) DEFAULT NULL,
  `gender` varchar(30) DEFAULT NULL,
  `date_of_birth` varchar(30) DEFAULT NULL,
  `ph_no` varchar(30) DEFAULT NULL,
  `email_id` varchar(30) DEFAULT NULL,
  `department` varchar(30) DEFAULT NULL,
  `post` varchar(30) DEFAULT NULL,
  `date_of_joining` varchar(30) DEFAULT NULL,
  `education_qualification` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`staff_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `staff_info`
--

INSERT INTO `staff_info` (`staff_id`, `first_name`, `last_name`, `address`, `gender`, `date_of_birth`, `ph_no`, `email_id`, `department`, `post`, `date_of_joining`, `education_qualification`) VALUES
(22, 'susan', 'bhujel', 'lamjung', 'male', '2052-03-05', '98463', 'bhujel.susan@yahoo.com', 'computer', 'student', '2069', 'no'),
(222, 'susan', 'bhujel', 'lamjung', 'male', '2052-03-05', '98463', 'bhujel.susan@yahoo.com', 'computer', 'student', '2069', 'no');

-- --------------------------------------------------------

--
-- Table structure for table `urine_test`
--

CREATE TABLE IF NOT EXISTS `urine_test` (
  `registration_id` int(11) DEFAULT NULL,
  `patient_name` varchar(50) DEFAULT NULL,
  `test_date` varchar(20) DEFAULT NULL,
  `age` int(20) DEFAULT NULL,
  `gender` varchar(50) DEFAULT NULL,
  `wbc` int(100) DEFAULT NULL,
  `rbc` int(100) DEFAULT NULL,
  `epthelial_cell` int(50) DEFAULT NULL,
  `sugar` varchar(50) DEFAULT NULL,
  `color` varchar(30) DEFAULT NULL,
  `reaction` varchar(20) DEFAULT NULL,
  KEY `registration_id` (`registration_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `urine_test`
--

INSERT INTO `urine_test` (`registration_id`, `patient_name`, `test_date`, `age`, `gender`, `wbc`, `rbc`, `epthelial_cell`, `sugar`, `color`, `reaction`) VALUES
(12, 'susan', '05-07-2015', 19, 'male', 5000, 4500, 3, 'mil', 'yellow', 'acidic');

--
-- Constraints for dumped tables
--

--
-- Constraints for table `appoinment`
--
ALTER TABLE `appoinment`
  ADD CONSTRAINT `appoinment_ibfk_1` FOREIGN KEY (`doctor_id`) REFERENCES `doctor_info` (`doctor_id`);

--
-- Constraints for table `billing`
--
ALTER TABLE `billing`
  ADD CONSTRAINT `billing_ibfk_1` FOREIGN KEY (`patient_id`) REFERENCES `patient_entry` (`patient_id`);

--
-- Constraints for table `discharge`
--
ALTER TABLE `discharge`
  ADD CONSTRAINT `discharge_ibfk_1` FOREIGN KEY (`patient_id`) REFERENCES `patient_entry` (`patient_id`);

--
-- Constraints for table `doctor_account`
--
ALTER TABLE `doctor_account`
  ADD CONSTRAINT `doctor_account_ibfk_1` FOREIGN KEY (`doctor_id`) REFERENCES `doctor_info` (`doctor_id`);

--
-- Constraints for table `staff_account`
--
ALTER TABLE `staff_account`
  ADD CONSTRAINT `staff_account_ibfk_1` FOREIGN KEY (`staff_id`) REFERENCES `staff_info` (`staff_id`);

--
-- Constraints for table `urine_test`
--
ALTER TABLE `urine_test`
  ADD CONSTRAINT `urine_test_ibfk_1` FOREIGN KEY (`registration_id`) REFERENCES `blood_test` (`registration_id`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
