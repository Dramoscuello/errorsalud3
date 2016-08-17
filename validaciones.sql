-- phpMyAdmin SQL Dump
-- version 4.5.2
-- http://www.phpmyadmin.net
--
-- Servidor: 127.0.0.1
-- Tiempo de generación: 17-08-2016 a las 01:39:13
-- Versión del servidor: 5.7.9
-- Versión de PHP: 5.6.16

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `validaciones`
--

DELIMITER $$
--
-- Procedimientos
--
DROP PROCEDURE IF EXISTS `Delete_tablas`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `Delete_tablas` ()  BEGIN
  TRUNCATE datosusutext;
  TRUNCATE datosusuerrores;
  TRUNCATE fichero;
END$$

DROP PROCEDURE IF EXISTS `Estado_Archivo`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `Estado_Archivo` ()  BEGIN
SELECT * FROM  (SELECT COUNT(e.td_t) AS 'Numero de Registros' FROM datosusutext e ) AS r INNER JOIN (select COUNT(d.error_) AS 'Registros Erroneos' FROM  datosusuerrores d)  AS nnn ;
  

END$$

DROP PROCEDURE IF EXISTS `PA_ceros`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_ceros` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Identificacion con ceros a la izquierda' from datosusutext where ni_t like '0%';
end$$

DROP PROCEDURE IF EXISTS `PA_HIJPAPELLIDO`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_HIJPAPELLIDO` ()  BEGIN
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t,       'Hij* en primer apellido' from datosusutext where papellido_t like 'HIJ%';
END$$

DROP PROCEDURE IF EXISTS `PA_HIJSAPELLIDO`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_HIJSAPELLIDO` ()  BEGIN 
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 	  'Hij* en segundo apellido' from datosusutext where sapellido_t like 'HIJ%';
END$$

DROP PROCEDURE IF EXISTS `PA_HIJSNOMBRE`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_HIJSNOMBRE` ()  BEGIN
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t,       'Hij* en segundo nombre' from datosusutext where snombre_t like 'HIJ%';
END$$

DROP PROCEDURE IF EXISTS `PA_identificacionerronea`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_identificacionerronea` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, mayor de edad con TI y menor con CC' from datosusutext where td_t='CC' and edad_t<18 union select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, mayor de edad con TI y menor con CC' from datosusutext where td_t='TI' and edad_t>=18;
end$$

DROP PROCEDURE IF EXISTS `PA_identificacionesrepetidas`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_identificacionesrepetidas` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, identificacion repetida' from datosusutext where ni_t in(select ni_t from datosusutext group by ni_t having count(ni_t)>1) order by ni_t;
end$$

DROP PROCEDURE IF EXISTS `PA_identificaciovacia`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_identificaciovacia` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, campo identificacion vacio' from datosusutext where ni_t is null;
end$$

DROP PROCEDURE IF EXISTS `PA_insertarfichero`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_insertarfichero` (IN `nm` VARCHAR(100))  BEGIN
	INSERT INTO fichero (name) VALUES(nm);
END$$

DROP PROCEDURE IF EXISTS `PA_logitudmax`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_logitudmax` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, campos sobrepasan longitud permitida' from datosusutext where character_length(td_t)>2 or character_length(ni_t)>20 or character_length(cea_t)>6 or character_length(tu_t)>1 or character_length(papellido_t)>30 or character_length(sapellido_t)>30 or character_length(pnombre_t)>20 or character_length(snombre_t)>20 or character_length(edad_t)>3 or character_length(ume_t)>1 or character_length(sexo_t)>1 or character_length(cdr_t)>2 or character_length(cmr_t)>3 or character_length(zrh_t)>1;
end$$

DROP PROCEDURE IF EXISTS `PA_papellidovacio`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_papellidovacio` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, primer apellido vacio' from datosusutext where papellido_t is null;
end$$

DROP PROCEDURE IF EXISTS `PA_pnombrevacio`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_pnombrevacio` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, primer nombre vacio' from datosusutext where pnombre_t is null;
end$$

DROP PROCEDURE IF EXISTS `PA_selectusuerrores`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_selectusuerrores` ()  BEGIN
	SELECT * FROM datosusuerrores;
END$$

DROP PROCEDURE IF EXISTS `PA_sexohombremal`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_sexohombremal` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error hombre con sexo de femenino' from datosusutext inner join sexomf on datosusutext.pnombre_t = sexomf.nombre and datosusutext.sexo_t <> sexomf.sexo and datosusutext.sexo_t = 'F';
end$$

DROP PROCEDURE IF EXISTS `PA_sexomujermal`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_sexomujermal` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error mujer con sexo masculino' from datosusutext inner join sexomf on datosusutext.pnombre_t = sexomf.nombre and datosusutext.sexo_t <> sexomf.sexo and datosusutext.sexo_t = 'M';
end$$

DROP PROCEDURE IF EXISTS `PA_sexonoexiste`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_sexonoexiste` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, sexo no existe' from datosusutext where sexo_t<>'M' and sexo_t<>'F';
end$$

DROP PROCEDURE IF EXISTS `PA_tipodocumentonoexiste`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_tipodocumentonoexiste` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, tipo de documento no existe' from datosusutext where td_t<>'CC' and td_t<>'TI' and td_t<>'CE' and td_t<>'PA' and td_t<>'RC' and td_t<>'AS' and td_t<>'MS' and td_t<>'MU';
end$$

DROP PROCEDURE IF EXISTS `PA_umenoexiste`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `PA_umenoexiste` ()  begin
	insert into datosusuerrores(td_, ni_, papellido_, sapellido_, pnombre_, snombre_, error_) select distinct td_t, ni_t, papellido_t, sapellido_t, pnombre_t, snombre_t, 'Error, unidad de medida de edad no existe' from datosusutext where ume_t<>1 and ume_t<>2 and ume_t<>3;
end$$

DROP PROCEDURE IF EXISTS `selectfichero`$$
CREATE DEFINER=`root`@`localhost` PROCEDURE `selectfichero` ()  BEGIN
	SELECT * FROM fichero;
END$$

DELIMITER ;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `datosusuerrores`
--

DROP TABLE IF EXISTS `datosusuerrores`;
CREATE TABLE IF NOT EXISTS `datosusuerrores` (
  `id_` int(255) NOT NULL AUTO_INCREMENT,
  `td_` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `ni_` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `papellido_` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `sapellido_` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `pnombre_` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `snombre_` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `error_` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  PRIMARY KEY (`id_`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_spanish_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `datosusutext`
--

DROP TABLE IF EXISTS `datosusutext`;
CREATE TABLE IF NOT EXISTS `datosusutext` (
  `td_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `ni_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `cea_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `tu_t` int(255) DEFAULT NULL,
  `papellido_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `sapellido_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `pnombre_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `snombre_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `edad_t` int(255) DEFAULT NULL,
  `ume_t` int(255) DEFAULT NULL,
  `sexo_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `cdr_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `cmr_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL,
  `zrh_t` varchar(255) COLLATE utf8_spanish_ci DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_spanish_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `fichero`
--

DROP TABLE IF EXISTS `fichero`;
CREATE TABLE IF NOT EXISTS `fichero` (
  `name` varchar(100) COLLATE utf8_spanish_ci NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_spanish_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `sexomf`
--

DROP TABLE IF EXISTS `sexomf`;
CREATE TABLE IF NOT EXISTS `sexomf` (
  `id` int(255) NOT NULL AUTO_INCREMENT,
  `nombre` varchar(20) COLLATE utf8_spanish_ci NOT NULL,
  `sexo` varchar(1) COLLATE utf8_spanish_ci NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=189 DEFAULT CHARSET=utf8 COLLATE=utf8_spanish_ci;

--
-- Volcado de datos para la tabla `sexomf`
--

INSERT INTO `sexomf` (`id`, `nombre`, `sexo`) VALUES
(1, 'DEIMER', 'M'),
(2, 'CECILIA', 'F'),
(3, 'JOSE', 'M'),
(4, 'LORENA', 'F'),
(5, 'LUIS', 'M'),
(6, 'CARLOS', 'M'),
(7, 'FERNANDO', 'M'),
(8, 'DIEGO', 'M'),
(9, 'MARLON', 'M'),
(10, 'FEDERICO', 'M'),
(11, 'JAVIER', 'M'),
(12, 'JAIDER', 'M'),
(13, 'MIGUEL', 'M'),
(14, 'JONATHAN', 'M'),
(15, 'EDUARDO', 'M'),
(16, 'YAIR', 'M'),
(17, 'HERNANDO', 'M'),
(18, 'ISAIAS', 'M'),
(19, 'PEDRO', 'M'),
(20, 'JEAN', 'M'),
(21, 'ANTONIO', 'M'),
(22, 'MARTIN', 'M'),
(23, 'ORLANDO', 'M'),
(24, 'OSCAR', 'M'),
(25, 'IVAN', 'M'),
(26, 'GUILLERMO', 'M'),
(27, 'SERGIO', 'M'),
(28, 'JESUS', 'M'),
(29, 'MARCOS', 'M'),
(30, 'GERONIMO', 'M'),
(31, 'ALEX', 'M'),
(32, 'OCTAVIO', 'M'),
(33, 'ISMAEL', 'M'),
(34, 'ROBERTO', 'M'),
(35, 'ENRIQUE', 'M'),
(36, 'MANUEL', 'M'),
(37, 'MARCIAL', 'M'),
(38, 'BARTOLOME', 'M'),
(39, 'TOMAS', 'M'),
(40, 'RICARDO', 'M'),
(41, 'LAZARO', 'M'),
(42, 'MOISES', 'M'),
(43, 'FELIX', 'M'),
(44, 'SEBASTIAN', 'M'),
(45, 'DANIEL', 'M'),
(46, 'PABLO', 'M'),
(47, 'DANIEL', 'M'),
(48, 'EUGENIO', 'M'),
(49, 'VIRGILIO', 'M'),
(50, 'EUSEBIO', 'M'),
(51, 'SAUL', 'M'),
(52, 'VICTOR', 'M'),
(53, 'EDGAR', 'M'),
(54, 'HECTOR', 'M'),
(55, 'ISAAC', 'M'),
(56, 'HORACIO', 'M'),
(57, 'EZEQUIEL', 'M'),
(58, 'KARINA', 'F'),
(59, 'LAURA', 'F'),
(60, 'TARCILIA', 'F'),
(61, 'PAOLA', 'F'),
(62, 'PAULA', 'F'),
(63, 'CAROLINA', 'F'),
(64, 'ELIZABETH', 'F'),
(65, 'SINDY', 'F'),
(66, 'OLGA', 'F'),
(67, 'KAREN', 'F'),
(68, 'PATRICIA', 'F'),
(69, 'GLADYS', 'F'),
(70, 'LILIANA', 'F'),
(71, 'JULIETH', 'F'),
(72, 'EYMIS', 'F'),
(73, 'GINA', 'F'),
(74, 'DANIELA', 'F'),
(75, 'VALENTINA', 'F'),
(76, 'CRISTINA', 'F'),
(77, 'YENNIFER', 'F'),
(78, 'JOHANA', 'F'),
(79, 'YURANIS', 'F'),
(80, 'NORMA', 'F'),
(81, 'KELLY', 'F'),
(82, 'ANA', 'F'),
(83, 'LINEYS', 'F'),
(84, 'DEICY', 'F'),
(85, 'EMILCE', 'F'),
(86, 'BRIGITH', 'F'),
(87, 'YONEDIS', 'F'),
(88, 'ASTRID', 'F'),
(89, 'DIANA', 'F'),
(90, 'LIBETH', 'F'),
(91, 'HELLA', 'F'),
(92, 'SANDY', 'F'),
(93, 'ESTHER', 'F'),
(94, 'ADELA', 'F'),
(95, 'ANGIE', 'F'),
(96, 'ANDRY', 'F'),
(97, 'DANIA', 'F'),
(98, 'EMILIS', 'F'),
(99, 'INES', 'F'),
(100, 'LEIDYS', 'F'),
(101, 'EVA', 'F'),
(102, 'LUISA', 'F'),
(103, 'IVONNE', 'F'),
(104, 'PIEDAD', 'F'),
(105, 'NATALIA', 'F'),
(106, 'VANESSA', 'F'),
(107, 'SUSANA', 'F'),
(108, 'ROSA', 'F'),
(109, 'MARIANA', 'F'),
(110, 'LUCIA', 'F'),
(111, 'ALFA', 'F'),
(112, 'MARELVIS', 'F'),
(113, 'CRISTIAN', 'M'),
(114, 'KARIN', 'F'),
(115, 'SILVIA', 'F'),
(116, 'MARIA', 'F'),
(117, 'DAVID', 'M'),
(118, 'JHON', 'M'),
(119, 'BERNARDO', 'M'),
(120, 'RAFAEL', 'M'),
(121, 'FERNANDO', 'M'),
(122, 'DIDIER', 'M'),
(123, 'DANILO', 'M'),
(124, 'ALVARO', 'M'),
(125, 'FELIPE', 'M'),
(126, 'GUSTAVO', 'M'),
(127, 'NEDER', 'M'),
(128, 'AARON', 'M'),
(129, 'ABEL', 'M'),
(130, 'JORGE', 'M'),
(131, 'ALEJANDRO', 'M'),
(132, 'CLAUDIA', 'M'),
(133, 'KIMBERLY', 'F'),
(134, 'SOFIA', 'F'),
(135, 'INDRID', 'F'),
(136, 'GABRIEL', 'M'),
(137, 'GABRIELA', 'F'),
(138, 'MERLYS', 'F'),
(139, 'LUZ', 'F'),
(140, 'MARCELA', 'F'),
(141, 'MARCELO', 'M'),
(142, 'WENDY', 'F'),
(143, 'LILIBETH', 'F'),
(144, 'MELISSA', 'F'),
(145, 'JULIANA', 'F'),
(146, 'JULIAN', 'M'),
(147, 'JULIO', 'M'),
(148, 'SAMIA', 'F'),
(149, 'ELVYS', 'M'),
(150, 'NELSON', 'M'),
(151, 'WILFER', 'M'),
(152, 'ADRIANA', 'F'),
(153, 'ADRIAN', 'M'),
(154, 'YERLIS', 'F'),
(155, 'MARA', 'F'),
(156, 'ARACELIS', 'F'),
(157, 'MAIRA', 'F'),
(158, 'VALERIA', 'F'),
(159, 'MILENA', 'F'),
(160, 'KATERINE', 'F'),
(161, 'THALIA', 'F'),
(162, 'ROMARIO', 'M'),
(163, 'FABIO', 'M'),
(164, 'FABIAN', 'M'),
(165, 'INDIRA', 'F'),
(166, 'NEILA', 'F'),
(167, 'LINA', 'F'),
(168, 'ISABELLA', 'F'),
(169, 'ANDRES', 'M'),
(170, 'ANDREA', 'F'),
(171, 'BLADIMIR', 'M'),
(172, 'ANDREINA', 'F'),
(173, 'BRIAN', 'M'),
(174, 'BRAYAN', 'M'),
(175, 'JAIDITH', 'F'),
(176, 'GEINER', 'M'),
(177, 'JEINER', 'M'),
(178, 'DUVAN', 'M'),
(179, 'MARTHA', 'F'),
(180, 'AMADA', 'F'),
(181, 'DEIVER', 'M'),
(182, 'CAMILO', 'M'),
(183, 'ENA', 'F'),
(184, 'EVER', 'M'),
(185, 'ALDAIR', 'M'),
(186, 'YILSON', 'M'),
(187, 'VIVIAN', 'F'),
(188, 'HUGO', 'M');

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
