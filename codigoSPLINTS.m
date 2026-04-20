clear; clc; close all;

% 1. LEER DATOS
fprintf('Esperando selección de archivo...\n');
[archivo, ruta] = uigetfile({'*.xlsx;*.xls'}, 'Seleccione su base de datos');
if isequal(archivo, 0), return; end

datos = readmatrix(fullfile(ruta, archivo));
x = datos(:, 1)'; 
y = datos(:, 2)';
n = length(x);

% 2. CREAR EL PARÁMETRO 't' (Del 1 hasta el número total de puntos)
t = 1:n; 

% 3. CREAR PUNTOS SUAVES PARA 't'
t_plot = linspace(1, n, 1000);

% 4. INTERPOLAR X e Y POR SEPARADO USANDO SPLINES (Lo mejor para la mano)
x_curva = spline(t, x, t_plot);
y_curva = spline(t, y, t_plot);

% --- GRÁFICA ---
figure('Color', 'w', 'Name', 'Silueta de la Mano');
plot(x_curva, y_curva, 'b-', 'LineWidth', 2, 'DisplayName', 'Curva Interpolada');
hold on;
plot(x, y, 'ro', 'MarkerFaceColor', 'r', 'DisplayName', 'Puntos Excel');

grid on; axis equal; % 'axis equal' asegura que no se deforme la mano
xlabel('Eje X'); ylabel('Eje Y');
title('Interpolación Paramétrica de Forma Cerrada');
legend('Location', 'best');