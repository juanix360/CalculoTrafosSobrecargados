%DESARROLLADOR: JUAN ORTIZ (2025)
%METODOLOGIA DE CALCULO: BRYAN ESTRELLA 
%DATA DE CALCULO: JENNY GUASHPA
%EMPRESA: CONSORCIO SIG-ELECTRIC
%DESCRIPCION: SCRIPT PARA CALCULAR LOS TRANSFORMADORES SOBRECARGADOS USANDO
%LA METODOLOGIA SUGERIDA POR BRYAN ESTRELLA EN DONDE EMBASE AL COSUMO SE
% ASIGNA UN ESTRATO QUE VA ENTRE LA A - E Y FINALMENTE SE ESCOGE EL ESTRATO
% MAS FRECUENTE Y SEGUN EL NUMERO DE USURIOS SE BUSCA EL CONSUMO EN LA
% TABLA DMD DE LA EMPRESA ELECTRICA QUITO. DENTRO DEL CALCULO SE TOMA
% ENCUENTA LA PERDIDA DEL 3.6% Y LA CAPACIDAD MAXIMA QUE PUEDE LLEGAR AL
% 1.25 DE LA NOMINAL.

%VERSION 3

clc
clear 
format LONGG

addpath('F_Funciones\');
inputFile='H:\Mi unidad\TrafosSobrecargados\';
addpath(inputFile);
folderOut='H:\Mi unidad\TrafosSobrecargados\Resultados\';
addpath(folderOut);
addpath('Z_TablasInput\');

%%                    CAMBIOS IMPORTANTES 
%Estos cambios se deben hacer en caso que se reemplace el archivo de
%ingreso o salida de datos, cambios de variable y hojas en excel

%%%%%%%%%%%%%%%%%%%%%%%Ingreso de datos%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Alimentador='-58';  
ExcelInputBase = [inputFile,'58_EL QUINCHE.xlsx'];

% BASE Hojas de lectura
sheetInBase1='TRAFO';
sheetInBase2='POSTE';
sheetInBase5='MEDIDORES';

ExcelTablaDMD='TablaDMD';

%%%%%%%%%%%%%%%%%%%%%% Salida de datos %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Obtener información del archivo
infoArchivo = dir(ExcelInputBase);

% Extraer la fecha de modificación
fechaModificacion = infoArchivo.date;
fechaModificacionDatetime = datetime(fechaModificacion, 'InputFormat', 'dd-MMM-yyyy HH:mm:ss');
fechaComoTexto = char(fechaModificacionDatetime);
soloFecha = extractBefore(fechaComoTexto, ' ');
ExcelOut= ['CargaTransformadores',Alimentador,'_', soloFecha,'.xlsx']; %Archivo de resultados
%Hojas de impresion
sheetOut1='CalculoCargaAnual';
sheetOut2='TransformadorDuplicados';
if ~exist(folderOut, 'dir')
    mkdir(folderOut); % Crear la carpeta si no existe
end

% sheetsOut = {sheetOut1};
% for i = 1:length(sheetsOut)
%     writecell({'Preparando archivo'}, fullfile(folderOut, ExcelOut), 'Sheet', sheetsOut{i}, 'Range', 'A1');
% end
% %%                   BORRADO DE DATOS 
% 
% % F_borrarDatosExcel(fullfile(folderOut, ExcelOut), sheetOut1);

%%                   LECTURA DE DATOS 
%%%%%%%%%%%%%%%%%%%%%%%%% TablaDMD %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
TablaDMD = readtable(ExcelTablaDMD, 'Sheet', 'Sheet1');
%%%%%%%%%%%%%%%%%%%%%%%%% BASE %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%++++++++++++++++Trafos++++++++++++++++++++++++++++++++++++++++++++++++++++
AlimTrafoBase = F_leerDatosExcel(ExcelInputBase, sheetInBase1); %Hoja completa de transformadores en la Base del Alimentador
%++++++++++++++++ Postes ++++++++++++++++++++++++++++++++++++++++++++++++++
PostesBase = F_leerDatosExcel(ExcelInputBase, sheetInBase2); %Hoja completa de medidores en la Base del Alimentador
%++++++++++++++++Medidores+++++++++++++++++++++++++++++++++++++++++++++++++
MedidoresBase = F_leerDatosExcel(ExcelInputBase, sheetInBase5); %Hoja completa de medidores en la Base del Alimentador
% Filtrar las filas donde la columna 'AGENCIA' no sea 'GRANDES CLIENT'
nInicial = height(MedidoresBase);
MedidoresBase = MedidoresBase(~strcmp(MedidoresBase.AGENCIA, 'GRANDES CLIENT'), :);
nFinal = height(MedidoresBase);

%%                      PROCESAMIENTO DE DATOS
%%%%%%%%%%%%%%%%%%%%%%%%%%% POSTES %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% AlimPostesCampo = renamevars(AlimPostesCampo, {'Seleccione el Poste', '','Seleccione la potencia de las luminarias'},{'PotenciaLuminarias'});
PostesBase= PostesBase (:, {'Codigo Puesto','POTENCIA'});
PostesBase = renamevars(PostesBase, {'Codigo Puesto','POTENCIA'},{'CodigoPuesto','POT_LUM'});

rowsWithEmptyValues = cellfun(@isempty, PostesBase.CodigoPuesto);
PostesBase(rowsWithEmptyValues, :) = [];

%--------------------------------------------------------------------------
% PostesBase.CANT_LUM = str2double(PostesBase.CANT_LUM);  % Convertir texto a números
% PostesBase.POT_LUM = str2double(PostesBase.POT_LUM);    % Asegurarte con ambas columnas
% 
% PostesBase.POT_TOTAL = PostesBase.CANT_LUM .* PostesBase.POT_LUM;
%--------------------------------------------------------------------------

%%%%%%%%%%%%%%%%%%%%%%%%%%% MEDIDORES %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
ResultMedidoresBase = MedidoresBase(:, {'CODIGOPUES','MDENUMEMP','CONSUMOP_1'});
ResultMedidoresBase = renamevars(ResultMedidoresBase, {'CODIGOPUES','CONSUMOP_1'},{'CodigoPuesto','CONSUMO'});

if iscell(ResultMedidoresBase.CONSUMO)
    % Convertir celdas a números, asignando NaN a valores no convertibles
    ResultMedidoresBase.CONSUMO = str2double(ResultMedidoresBase.CONSUMO);
end

% Encontrar valores únicos y sus índices
codigos = AlimTrafoBase.("Codigo Puesto"); 
[unicos, ~, idx] = unique(codigos, 'stable');
conteo = histcounts(idx, numel(unicos));
% Identificar valores duplicados (ocurrencias > 1)
valoresDuplicados = unicos(conteo > 1);
tablaDuplicados = AlimTrafoBase(ismember(codigos, valoresDuplicados), :); % Contiene solo los duplicados

% Crear nuevas tablas
TrafoAlimentador = AlimTrafoBase.("Codigo Puesto")(~ismember(codigos, valoresDuplicados), :);  % Mantiene solo valores únicos

% Categorización de consumos
categorias = ["A1", "A", "B", "C", "D", "E"];

%++++++++++++++++++++++Calculo Estratos +++++++++++++++++++++++++++++++++++

numTrafos = length(TrafoAlimentador);
% Preasignar la tabla con las columnas necesarias
Resultados = table('Size', [numTrafos, 5], ...
    'VariableTypes', {'string','double', 'double', 'double','double'}, ...
    'VariableNames', {'CODIGOPUES', 'PotenciaNominal', 'CapacidadMaxima','DemandaConsumo','CapacidadCalculada'});

for i = 1:length(TrafoAlimentador)
    % Filtrar medidores correspondientes al transformador actual
    TrafoActual = TrafoAlimentador(i);
    MedidoresXTrafo = ResultMedidoresBase(strcmp(ResultMedidoresBase.CodigoPuesto, TrafoActual), :); % Filtra por índice

    % Agregar una nueva columna 'Categoria' en la tabla MedidoresXTrafo
    MedidoresXTrafo.Categoria = strings(height(MedidoresXTrafo), 1); % Inicializar columna como cadenas

    % Categorización de consumos usando límites
    MedidoresXTrafo.Categoria = strings(height(MedidoresXTrafo), 1); % Inicializar columna

        MedidoresXTrafo.Categoria(MedidoresXTrafo.CONSUMO <= 100) = "E";
        MedidoresXTrafo.Categoria(MedidoresXTrafo.CONSUMO > 100 & MedidoresXTrafo.CONSUMO <= 150) = "D";
        MedidoresXTrafo.Categoria(MedidoresXTrafo.CONSUMO > 150 & MedidoresXTrafo.CONSUMO <= 250) = "C";
        MedidoresXTrafo.Categoria(MedidoresXTrafo.CONSUMO > 250 & MedidoresXTrafo.CONSUMO <= 350) = "B";
        MedidoresXTrafo.Categoria(MedidoresXTrafo.CONSUMO > 350 & MedidoresXTrafo.CONSUMO <= 500) = "A";
        MedidoresXTrafo.Categoria(MedidoresXTrafo.CONSUMO > 500) = "A1";

    % Calcular la frecuencia de cada categoría
    Frecuencias = arrayfun(@(cat) sum(MedidoresXTrafo.Categoria == cat), categorias)';
    
    CatFrecuencia = table(categorias', Frecuencias, 'VariableNames', {'Categorias', 'Frecuencias'});
    % ResultadoCatFrecuencia(:,2) = str2double(ResultadoCatFrecuencia(:,2));
    % Paso 1: Encontrar el valor máximo en G
    max_value = max(CatFrecuencia.Frecuencias);
    % Paso 2: Encontrar la posición de ese valor máximo en G
   index = find(CatFrecuencia.Frecuencias == max_value);

   % Paso 3: Usar el índice para obtener el valor correspondiente de H
   UsuarioMasComun = CatFrecuencia.Categorias{index};  % El valor correspondiente en la columna H
   sumaUsuarios = sum(CatFrecuencia.Frecuencias);
   idxDMD = find(TablaDMD.N_usuarios >= sumaUsuarios, 1);
   DemandaTotal = TablaDMD{idxDMD, UsuarioMasComun};
%++++++++++++++++++++++Calculo Postes Potencia+++++++++++++++++++++++++++++
    
   % Filtrar la tabla para obtener solo las filas con el transformador actual
   filasFiltradas = PostesBase(strcmp(PostesBase.CodigoPuesto, TrafoActual), :);

   % Sumar las potencias de las luminarias para ese transformador
   sumaPotenciaPostes = sum(cellfun(@str2double, filasFiltradas.POT_LUM));
    
    %----------------------------------------------------------------------
    % indices = strcmp(PostesBase.CodigoPuesto, TrafoActual);
    %     % Sumar los valores de PLuminariasNumeros del grupo actual
    %     % Agrupar por la columna TRAFO y sumar POT_LUM
    % sumaPotenciaPostes = sum(PostesBase.POT_TOTAL(indices));
    %----------------------------------------------------------------------
        % Calcular el resultado de la operación suma / (1000 * 0.92)
    PotenciaPostes = sumaPotenciaPostes / (1000 * 0.92);

    CapacidadCalculadaTrafo =DemandaTotal+PotenciaPostes;

    CapacidadCalculadaTrafo =CapacidadCalculadaTrafo + CapacidadCalculadaTrafo*0.036; 

   % Filtrar el índice donde "CODIGOPUES" coincide con "TrafoActual"
   indiceTrafo = strcmp(AlimTrafoBase.("Codigo Puesto"), TrafoActual); % Si es texto
    % Verificar si se encontró el transformador en la tabla
    if any(indiceTrafo)
        % Extraer el valor de la potencia de la columna "POTENCIAKV"
        PotenciaNominalTrafoActual = AlimTrafoBase.("Potencia (kva)")(indiceTrafo);
    else
        % Manejar el caso donde no se encontró el transformador
        PotenciaNominalTrafoActual = NaN; % Asignar NaN o un valor predeterminado
        warning(['El TrafoActual ', char(TrafoActual), ' no se encontró en AlimTrafoBase.']);
    end

     % Guardar los resultados en la tabla preasignada
     Resultados.CODIGOPUES(i) = TrafoActual;
  
     Resultados.PotenciaNominal(i) = str2double(PotenciaNominalTrafoActual);
     Resultados.CapacidadMaxima(i) = Resultados.PotenciaNominal(i)*1.25;
     Resultados.DemandaConsumo(i) = DemandaTotal;
     Resultados.CapacidadCalculada(i) = CapacidadCalculadaTrafo;

     disp(['Procesando transformador ', TrafoActual]);

end

% Asegurarse de que PotenciaNominal sea de tipo numérico
Resultados.Estado = repmat("OK", height(Resultados), 1); % Inicializar con "OK"
sobrecargados = Resultados.CapacidadCalculada > Resultados.CapacidadMaxima;
Resultados.Estado(sobrecargados) = "SOBRECARGADO";

%%%%%%%%%%%%%%%%%%%% Impresion Resultados %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
rangePrintFirst=1; 
rangeMedidoresAlert=strcat('A', num2str(rangePrintFirst));
writetable(Resultados,fullfile(folderOut, ExcelOut),'Sheet', sheetOut1,'Range',rangeMedidoresAlert);%Imprime tabla de Alertas
writetable(tablaDuplicados,fullfile(folderOut, ExcelOut),'Sheet', sheetOut2,'Range',rangeMedidoresAlert);%Imprime tabla de Alertas

%%                          GRAFICAS

% % Calcular porcentaje de carga
% Resultados.PorcentajeCarga = (Resultados.CapacidadCalculada ./ Resultados.CapacidadMaxima) * 100;
% 
% % Agrupar transformadores en rangos de porcentaje de carga
% edges = [0, 25, 50, 75, 100, Inf]; % Rango de porcentaje de carga
% labels = {'0-25%', '25-50%', '50-75%', '75-100%', '>100%'};
% gruposCarga = discretize(Resultados.PorcentajeCarga, edges, 'categorical', labels);
% 
% % Contar transformadores en cada rango
% conteoGrupos = countcats(gruposCarga);
% 
% % Graficar
% figure;
% bar(categorical(labels), conteoGrupos);
% title('Distribución de Transformadores por Rango de Carga Maxima');
% xlabel('Rango de Carga (%)');
% ylabel('Número de Transformadores');
% grid on;