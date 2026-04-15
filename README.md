<h1><img width="25" height="25" alt="Window Icon" src="https://github.com/user-attachments/assets/c760c12a-5ca5-4a7a-8f94-59a1e28457dc" /> Mariser!</h1>
<h2>💾 "Instalación"</h2>
Instalar el programa para útilizarlo es muy sencillo siguiendo estos pasos:<br>
<ol>
    <li><b>Haz click sobre "Mariser! Última Versión.exe"</b></li>
    <img width="1920" height="1080" alt="Clicka la última versión 2" src="https://github.com/user-attachments/assets/850e25c5-2b20-4a25-b949-cc61f486202e" />
    <br><br>
    <li><b>Haz click sobre "Raw"</b></li>
    <img width="1920" height="1080" alt="Clickea Raw" src="https://github.com/user-attachments/assets/b711177d-b370-4a2d-a636-aa32c1cef0c6" />
    <br><br>
    <li><b>Haz click sobre el botón de Descargas</b></li>
    <img width="1920" height="1080" alt="Click download" src="https://github.com/user-attachments/assets/73facfe3-2652-4ba8-bf18-8c37edc7dece" />
    <br><br>
    <li><b>Haz click sobre la descarga para abrir el programa</b></li>
    <img width="1920" height="1080" alt="Click Mariser" src="https://github.com/user-attachments/assets/ec3e47f7-892f-4b75-acd5-9a3acf995f12" />
    <br><br>
</ol>
<h3>Opcional: Poner acceso rápido</h3>
Si bien de esa manera puedes útilizar el programa, es inconveniente para darle uso regularmente...<br>
Para solucionar esto, podemos hacer...<br>

<ol>
   <li><b>Repite el paso número 3</b></li>
   <img width="1920" height="1080" alt="Click download" src="https://github.com/user-attachments/assets/73facfe3-2652-4ba8-bf18-8c37edc7dece" />
   <br><br>
   <li><b>Pon el ratón sobre la descarga, y haz click sobre el icono de la carpeta</b></li>
   <img width="1920" height="1080" alt="Hover The Mouse and Click over file icon 3" src="https://github.com/user-attachments/assets/00432438-e260-4c98-8284-78be49aab8db" />
   <br><br>

   <li><b>Haz click derecho sobre el archivo resaltado en azul, busca "Anclar a la barra de tareas" y dale click</b></li>
   <img width="1920" height="1080" alt="Right click and &#39;anclar a la barra de tareas&#39;" src="https://github.com/user-attachments/assets/92e12aab-c423-4f64-ad85-828ff963b0b5" />
   <strong>Nota:</strong> Dependiendo de la versión del sistema, puede ser que esta opción aparezca al hacer click a un botón que diga "Más opciones".
   <br><br>
</ol>
Esto hace que puedas abrir el programa rápidamente cuando lo neceistes simplemente haciendo click en el icono en la barra de tareas.
<img width="1920" height="1080" alt="Click Mariser Barra Tareas" src="https://github.com/user-attachments/assets/a7a3bb61-d275-4c46-93c8-eb8e5bab102d" />
<br><br>
<h3>Instalar desde código fuente <b>(avanzado)</b></h3>
Esta opción está pensada exclusivamente para desarrolladores, por tal motivo, esta sección tiene un tono más técnico, y asume que tienes una idea de lo que estás haciendo...
<ol>
   <li><b>Clona el repositorio a tu máquina</b></li>
   <li><b>En caso de que no lo tengas ya, instala python</b></li>
       En sistemas windows modernos(win 10 y superior) puedes útilizar winget para instalar python.<br>
       En tu interprete de comandos favorito, ejecuta: <code>winget install python</code><br>
       Si tienes el gestor de paquetes chocolately, también puedes instalarlo por esa vía: <code>choco install python</code><br>
       Si no dispones de uno ni de otro, o por algún otro motivo puedes conseguir python a traves de binarios precompilados: <a href=https://www.python.org/downloads/>Página de descargas de python.org</a>
   <li><b>Instala pyinstaller</b></li>
       Con python ya instalado, deberías ser capaz de utilizar pip, e instalar pyinstaller.<br>
       Ejecuta: <code>pip install pyinstaller</code><br>
       <b>Nota: </b>en caso de que el sistema no reconozca pip como un comando, probablemente sea el caso de que no vino instalado con tu versión de python, o que hay un problema en el entorno de tu OS. En cualquier caso, buena suerte con eso :)
    <li><b>Construye el ejecutable</b></li>
        Primero, dirigete al directorio raíz del repositorio en tu máquina. Luego, deberás ejecutar el batch script úbicado en /Releasing. El cuál tiene el comando exacto para reconstruir el archivo ejecutable a partir del script de python.<br>
        Para eso, de nuevo, una vez en el directorio raíz del repositorio, ejecuta: <code>"Releasing\Pyinstaller command.bat"</code><br>
        Esto generará el archivo ejecutable en tu directorio actual. Claramente podrías ejecutar el script desde cualquier lugar que te plazca, y/o moverlo a voluntad, pero esta manera es la más fácil de enseñar.<br>
</ol>
<h2>ℹ️ Acerca del programa</h2>
Mariser! es un programa que automatiza el checkeo de lotes contables. Es un programa creado por Lucas Da Silva en 2026 como regalo para el equipo De La Plaza Hotel.<br>
A partir del archivo de Mayores Contables generado automaticamente por el sistema de hoteleria, genera un resumen con el estado de cada lote en el archivo.<br>
<br>
Es un programa sencillo, pensado para ser una herramienta extremadamente fácil de usar sin ningún tipo de conocimiento técnico.<br>
Pensada para fácilitar al máximo un trabajo que es largo, arduo, tedioso, extremadamente importante y propenso al error humano.<br>
<br>
