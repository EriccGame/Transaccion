USE Temporales

CREATE TABLE datos_empleados
(
	id INT,
	nombre VARCHAR(MAX),
	direccion VARCHAR(MAX)
)

CREATE TYPE type_empleados AS TABLE 
(
	id INT PRIMARY KEY,
	nombre VARCHAR(MAX),
	direccion VARCHAR(MAX)
)

CREATE PROCEDURE spRegistrarEmpleado
	@datos_empleados type_empleados READONLY
AS
BEGIN

	DECLARE @id INT = 0
	DECLARE @nombre VARCHAR(MAX) = ''
	DECLARE @direccion VARCHAR(MAX) = ''

	SELECT @id = id, @nombre = nombre, @direccion = direccion
	FROM @datos_empleados
	
	INSERT INTO datos_empleados VALUES (@id, @nombre, @direccion)

END
GO

SELECT * FROM datos_empleados