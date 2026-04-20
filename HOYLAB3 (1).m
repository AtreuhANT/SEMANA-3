clear; clc; close all;

% 1. SELECCIÓN DEL ARCHIVO
fprintf('Seleccione el archivo de Excel con los datos...\n');
[archivo, ruta] = uigetfile({'*.xlsx;*.xls', 'Archivos de Excel (*.xlsx, *.xls)'}, 'Seleccione el archivo de datos');

if isequal(archivo, 0)
    disp('Usuario canceló la selección.');
    return;
else
    fullPath = fullfile(ruta, archivo);
    % Leemos la tabla (asume que x está en la col 1 y y en la col 2)
    datos = readmatrix(fullPath);
    x = datos(:, 1)'; % Extrae columna 1 y la transpone a fila
    y = datos(:, 2)'; % Extrae columna 2 y la transpone a fila
end

% 2. CONFIGURACIÓN
n = length(x);
xi = 3.5; % Punto a evaluar
x_plot = linspace(min(x), max(x), 500);

% --- MÉTODO DE LAGRANGE ---
y_lagrange = zeros(size(x_plot));
for k = 1:length(x_plot)
    L = ones(1, n);
    for i = 1:n
        for j = 1:n
            if i ~= j
                L(i) = L(i) * (x_plot(k) - x(j)) / (x(i) - x(j));
            end
        end
    end
    y_lagrange(k) = sum(y .* L);
end

% --- MÉTODO DE NEWTON ---
D = zeros(n, n);
D(:,1) = y'; 
for j = 2:n
    for i = 1:n-j+1
        D(i,j) = (D(i+1,j-1) - D(i,j-1)) / (x(i+j-1) - x(i));
    end
end
coef_newton = D(1,:);
y_newton = zeros(size(x_plot));
for k = 1:length(x_plot)
    y_newton(k) = coef_newton(1);
    prod_term = 1;
    for i = 2:n
        prod_term = prod_term * (x_plot(k) - x(i-1));
        y_newton(k) = y_newton(k) + coef_newton(i) * prod_term;
    end
end

% 3. EVALUACIÓN EN EL PUNTO xi
val_L = 0; % Re-calculando rápido para el punto exacto
L_pt = ones(1,n);
for i=1:n
    for j=1:n
        if i~=j, L_pt(i) = L_pt(i)*(xi-x(j))/(x(i)-x(j)); end
    end
end
val_lagrange = sum(y .* L_pt);
% --- COMPARACIÓN DE RESULTADOS ---
% Evaluamos Newton en el punto xi de forma específica
val_newton = coef_newton(1);
prod_term_xi = 1;
for i = 2:n
    prod_term_xi = prod_term_xi * (xi - x(i-1));
    val_newton = val_newton + coef_newton(i) * prod_term_xi;
end

% Mostrar tabla de comparación
fprintf('\n=========================================');
fprintf('\n       RESULTADOS DE INTERPOLACIÓN       ');
fprintf('\n=========================================');
fprintf('\n Punto evaluado (x): %.2f', xi);
fprintf('\n Método de Lagrange: %.6f', val_lagrange);
fprintf('\n Método de Newton:   %.6f', val_newton);
fprintf('\n Diferencia:         %.2e', abs(val_lagrange - val_newton));
fprintf('\n=========================================\n');

% 4. GRÁFICA Y RESULTADOS
fprintf('\nArchivo cargado: %s\n', archivo);
fprintf('Interpolación de grado: %d\n', n-1);
fprintf('Evaluación en x = %.1f: %.4f\n', xi, val_lagrange);

figure('Color', 'w');
plot(x_plot, y_lagrange, 'b-', 'LineWidth', 2, 'DisplayName', 'Curva Interpolada'); hold on;
plot(x, y, 'ro', 'MarkerFaceColor', 'r', 'DisplayName', 'Datos Excel');
plot(xi, val_lagrange, 'ks', 'MarkerFaceColor', 'y', 'MarkerSize', 10, 'DisplayName', 'Punto 3.5');
grid on; legend('Location', 'best');
title(['Interpolación desde: ', archivo]);