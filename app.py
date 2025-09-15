import openpyxl
from io import BytesIO
import datetime
from flask import Flask, request, jsonify, render_template_string, session, send_file

# Configuración de la aplicación Flask
app = Flask(__name__)
app.secret_key = 'super_secreto_y_seguro' # Cambia esto por una clave secreta real en producción

# La plantilla HTML, incrustada en el código
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es" class="h-full">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestor de Finanzas Personales</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        body {
            font-family: 'Inter', sans-serif;
            transition: background-color 0.3s ease;
        }
        .sidebar-button {
            @apply w-full text-left py-3 px-4 my-2 rounded-xl transition-all duration-200 ease-in-out font-medium;
        }
        .sidebar-button.active {
            @apply bg-indigo-600 text-white shadow-lg transform scale-105;
        }
        .sidebar-button:hover:not(.active) {
            @apply bg-gray-700 text-white shadow-md;
        }
        .content-section {
            display: none;
            animation: fadeIn 0.5s ease-in-out;
        }
        .content-section.active {
            display: block;
        }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .input-group {
            @apply flex flex-col mb-4;
        }
        .input-group label {
            @apply text-sm font-semibold mb-1 text-white ;
        }
       .input-group input, .input-group select {
        @apply p-3 rounded-xl text-white border border-gray-600 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent;
        background-color: transparent; /* Fondo transparente */
        }
        .icon {
            @apply inline-block mr-2;
        }
        .btn-action {
            @apply py-2 px-4 rounded-xl text-sm font-semibold transition-colors duration-200 ease-in-out;
        }
        .summary-card {
            @apply bg-gray-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between items-start transition-transform duration-300 hover:scale-[1.01];
        }
    </style>
</head>
<body class="bg-gray-950 text-white min-h-screen flex flex-col md:flex-row">
    <!-- Sidebar -->
    <aside class="w-full md:w-64 bg-gray-900 p-6 shadow-2xl rounded-xl m-4 md:mr-0">
        <h1 class="text-3xl font-bold mb-8 text-center text-indigo-400">FinanzasApp</h1>
        <nav>
            <button class="bg-blue-600 hover:bg-green-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full" onclick="showSection('panel')">
                <span class="icon">📊</span> Panel Principal
            </button>
            <button class="bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full" onclick="showSection('ingresos')">
                <span class="icon">➕</span> Ingresos
            </button>
            <button class="bg-red-600 hover:bg-green-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full" onclick="showSection('egresos')">
                <span class="icon">➖</span> Egresos
            </button>
            <button class="bg-yellow-600 hover:bg-green-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full" onclick="showSection('ahorros')">
                <span class="icon">💰</span> Ahorros
            </button>
            <button class="bg-purple-600 hover:bg-green-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full" onclick="showSection('deudas')">
                <span class="icon">📑</span> Deudas
            </button>
            <button class="bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full" onclick="showSection('por-cobrar')">
                <span class="icon">📌</span> Por Cobrar
            </button>
            <a href="/exportar" class="bg-gray-300 hover:bg-gray-400 text-gray-800 font-bold py-2 px-4 mt-4 rounded inline-flex items-center" target="_blank" onclick="showModal('Exportación', 'Los datos se están exportando a un archivo Excel. La descarga comenzará pronto.');">
                <span class="icon">📤</span> Exportar Excel
            </a>
        </nav>
    </aside>

    <!-- Main Content -->
    <main class="flex-1 bg-gray-900 p-8 shadow-2xl rounded-xl m-4">
        <!-- Panel Principal -->
        <section id="panel" class="content-section active">
            <h2 class="text-3xl font-bold mb-6 text-indigo-400">📊 Resumen Financiero</h2>
            <div id="panel-data" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                <!-- Data will be populated by JS -->
                <p>Cargando datos...</p>
            </div>
        </section>

        <!-- Ingresos Section -->
        <section id="ingresos" class="content-section">
            <h2 class="text-3xl font-bold mb-6 text-green-400">➕ Añadir Ingreso</h2>
            <form id="ingresos-form" class="bg-gray-800 p-6 rounded-xl shadow-inner">
                <div class="input-group">
                    <label for="monto-ingreso">Monto:</label>
                    <input type="number" id="monto-ingreso" name="monto" placeholder="Ej: 1500" required>
                </div>
                <div class="input-group">
                    <label for="procedencia">Procedencia:</label>
                    <input type="text" id="procedencia" name="procedencia" placeholder="Ej: Salario" required>
                </div>
                <div class="input-group">
                    <label for="metodo-ingreso">Método:</label>
                    <select id="metodo-ingreso" name="metodo" class="w-full" onchange="toggleDeudorInput()">
                        <option value="Efectivo">Efectivo</option>
                        <option value="Transferencia">Transferencia</option>
                        <option value="Crédito">Crédito</option>
                    </select>
                </div>
                <div class="input-group" id="deudor-group" style="display: none;">
                    <label for="deudor">¿Quién debe?:</label>
                    <input type="text" id="deudor" name="deudor" placeholder="Ej: Juan">
                </div>
                <button type="submit" class="bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full">Guardar Ingreso</button>
            </form>
        </section>

        <!-- Egresos Section -->
        <section id="egresos" class="content-section">
            <h2 class="text-3xl font-bold mb-6 text-red-400">➖ Añadir Egreso</h2>
            <form id="egresos-form" class="bg-gray-800 p-6 rounded-xl shadow-inner">
                <div class="input-group">
                    <label for="monto-egreso">Monto:</label>
                    <input type="number" id="monto-egreso" name="monto" placeholder="Ej: 500" required>
                </div>
                <div class="input-group">
                    <label for="concepto">Concepto de gasto:</label>
                    <input type="text" id="concepto" name="concepto" placeholder="Ej: Supermercado" required>
                </div>
                <div class="input-group">
                    <label for="metodo-egreso">Método:</label>
                    <select id="metodo-egreso" name="metodo" class="w-full">
                        <option value="Efectivo">Efectivo</option>
                        <option value="Transferencia">Transferencia</option>
                        <option value="Crédito">Crédito</option>
                    </select>
                </div>
                <div class="input-group">
                    <label for="persona-egreso">Pagar deuda a:</label>
                    <select id="persona-egreso" name="persona" class="w-full">
                        <option value="Nadie">Nadie</option>
                    </select>
                </div>
                <button type="submit" class="bg-red-600 hover:bg-red-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full">Guardar Egreso</button>
            </form>
        </section>

        <!-- Ahorros Section -->
        <section id="ahorros" class="content-section">
            <h2 class="text-3xl font-bold mb-6 text-yellow-400">💰 Añadir Ahorro</h2>
            <form id="ahorros-form" class="bg-gray-800 p-6 rounded-xl shadow-inner">
                <div class="input-group">
                    <label for="monto-ahorro">Monto a ahorrar:</label>
                    <input type="number" id="monto-ahorro" name="monto" placeholder="Ej: 200" required>
                </div>
                <button type="submit" class="bg-yellow-600 hover:bg-yellow-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full">Guardar Ahorro</button>
            </form>
        </section>

        <!-- Deudas Section -->
        <section id="deudas" class="content-section">
            <h2 class="text-3xl font-bold mb-6 text-purple-400">📑 Deudas (Yo debo)</h2>
            <div id="deudas-list" class="bg-gray-800 p-6 rounded-xl shadow-inner mb-6">
                <p>Cargando deudas...</p>
            </div>
            <h3 class="text-xl font-bold mt-8 mb-4">➕ Añadir nueva deuda</h3>
            <form id="deudas-form" class="bg-gray-800 p-6 rounded-xl shadow-inner">
                <div class="input-group">
                    <label for="persona-deuda">Nombre de la persona:</label>
                    <input type="text" id="persona-deuda" name="persona" placeholder="Ej: Sofía" required>
                </div>
                <div class="input-group">
                    <label for="monto-deuda">Monto de la deuda:</label>
                    <input type="number" id="monto-deuda" name="monto" placeholder="Ej: 100" required>
                </div>
                <button type="submit" class="bg-purple-600 hover:bg-purple-700 text-white font-bold py-3 px-6 rounded-xl mt-4 w-full">Guardar Deuda</button>
            </form>
        </section>

        <!-- Por Cobrar Section -->
        <section id="por-cobrar" class="content-section">
            <h2 class="text-3xl font-bold mb-6 text-teal-400">📌 Por Cobrar (Me deben)</h2>
            <div id="por-cobrar-list" class="bg-gray-800 p-6 rounded-xl shadow-inner">
                <p>Cargando cuentas por cobrar...</p>
            </div>
        </section>
    </main>

    <!-- Modal for messages -->
    <div id="modal-container" class="fixed inset-0 bg-black bg-opacity-75 flex items-center justify-center z-50 transition-opacity duration-300 opacity-0 pointer-events-none">
        <div class="bg-gray-800 p-6 rounded-xl shadow-2xl max-w-sm w-full transform transition-transform duration-300 scale-95 border border-gray-700">
            <div class="flex justify-between items-center mb-4">
                <h3 id="modal-title" class="text-xl font-bold text-indigo-400"></h3>
                <button onclick="closeModal()" class="text-gray-400 hover:text-white">&times;</button>
            </div>
            <p id="modal-message" class="text-gray-300"></p>
            <div class="mt-6 text-right">
                <button onclick="closeModal()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 px-4 rounded-lg">Cerrar</button>
            </div>
        </div>
    </div>

    <script>
        // Funciones del Frontend
        let activeSection = 'panel';

        function showSection(sectionId) {
            // Actualiza la sección activa
            const sections = document.querySelectorAll('.content-section');
            sections.forEach(sec => sec.classList.remove('active'));
            document.getElementById(sectionId).classList.add('active');

            // Actualiza el botón activo en la barra lateral
            const buttons = document.querySelectorAll('.sidebar-button');
            buttons.forEach(btn => btn.classList.remove('active'));
            document.querySelector(`[onclick="showSection('${sectionId}')"]`).classList.add('active');
            
            // Recargar datos para las secciones que lo necesitan
            if (['panel', 'deudas', 'por-cobrar', 'egresos'].includes(sectionId)) {
                fetchAndRenderData();
            }
        }

        async function fetchAndRenderData() {
            try {
                const response = await fetch('/get_data');
                const data = await response.json();
                renderPanel(data);
                renderDeudas(data);
                renderPorCobrar(data);
                populateEgresosDeudas(data);
            } catch (error) {
                console.error('Error al obtener los datos:', error);
            }
        }

        function renderPanel(data) {
            const panelData = document.getElementById('panel-data');
            const totalIngresos = data.ingresos.reduce((sum, item) => sum + item.monto, 0);
            const totalEgresos = data.egresos.reduce((sum, item) => sum + item.monto, 0);
            const totalAhorros = data.ahorros.reduce((sum, item) => sum + item, 0);
            const numDeudas = Object.keys(data.deudas).length;
            const numPorCobrar = Object.keys(data.por_cobrar).length;

            panelData.innerHTML = `
                <div class="summary-card">
                    <div class="text-4xl">💵</div>
                    <p class="text-lg font-medium text-gray-400 mt-2">Capital Disponible</p>
                    <p class="text-3xl font-bold text-green-400 mt-1">$${data.capital.toFixed(2)}</p>
                </div>
                <div class="summary-card">
                    <div class="text-4xl">💸</div>
                    <p class="text-lg font-medium text-gray-400 mt-2">Ingresos Totales</p>
                    <p class="text-3xl font-bold text-white mt-1">$${totalIngresos.toFixed(2)}</p>
                </div>
                <div class="summary-card">
                    <div class="text-4xl">📉</div>
                    <p class="text-lg font-medium text-gray-400 mt-2">Egresos Totales</p>
                    <p class="text-3xl font-bold text-white mt-1">$${totalEgresos.toFixed(2)}</p>
                </div>
                <div class="summary-card">
                    <div class="text-4xl">💰</div>
                    <p class="text-lg font-medium text-gray-400 mt-2">Ahorros</p>
                    <p class="text-3xl font-bold text-yellow-400 mt-1">$${totalAhorros.toFixed(2)}</p>
                </div>
                <div class="summary-card">
                    <div class="text-4xl">🤝</div>
                    <p class="text-lg font-medium text-gray-400 mt-2">Deudas</p>
                    <p class="text-3xl font-bold text-purple-400 mt-1">${numDeudas} pendientes</p>
                </div>
                <div class="summary-card">
                    <div class="text-4xl">📈</div>
                    <p class="text-lg font-medium text-gray-400 mt-2">Por Cobrar</p>
                    <p class="text-3xl font-bold text-teal-400 mt-1">${numPorCobrar} pendientes</p>
                </div>
            `;
        }

        function renderDeudas(data) {
            const deudasList = document.getElementById('deudas-list');
            deudasList.innerHTML = '';
            if (Object.keys(data.deudas).length === 0) {
                deudasList.innerHTML = '<p class="text-gray-400">✅ No tienes deudas registradas.</p>';
            } else {
                for (const [persona, monto] of Object.entries(data.deudas)) {
                    deudasList.innerHTML += `<div class="p-4 bg-gray-700 rounded-lg my-2 flex justify-between items-center shadow-md"><span class="font-semibold">${persona}:</span> <span class="text-xl text-red-400">$${monto.toFixed(2)}</span></div>`;
                }
            }
        }

        function renderPorCobrar(data) {
            const porCobrarList = document.getElementById('por-cobrar-list');
            porCobrarList.innerHTML = '';
            if (Object.keys(data.por_cobrar).length === 0) {
                porCobrarList.innerHTML = '<p class="text-gray-400">✅ No tienes cuentas por cobrar.</p>';
            } else {
                for (const [persona, monto] of Object.entries(data.por_cobrar)) {
                    const div = document.createElement('div');
                    div.className = 'p-4 bg-gray-700 rounded-lg my-2 flex flex-col sm:flex-row justify-between items-start sm:items-center space-y-2 sm:space-y-0 shadow-md';
                    div.innerHTML = `
                        <span class="font-semibold text-xl">${persona}: <span class="text-teal-400">$${monto.toFixed(2)}</span></span>
                        <div class="flex space-x-2">
                            <button class="btn-action bg-blue-500 hover:bg-blue-600 text-white" onclick="editPorCobrar('${persona}', ${monto})">✏️ Editar</button>
                            <button class="btn-action bg-green-500 hover:bg-green-600 text-white" onclick="deletePorCobrar('${persona}')">✅ Pagado</button>
                        </div>
                    `;
                    porCobrarList.appendChild(div);
                }
            }
        }

        async function editPorCobrar(persona, currentMonto) {
            // Reemplazado por un modal personalizado
            const newMonto = await showCustomPrompt(`Ingrese el nuevo monto para ${persona}:`);
            if (newMonto !== null && !isNaN(parseFloat(newMonto))) {
                const response = await fetch('/editar_por_cobrar', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ persona, monto: parseFloat(newMonto) })
                });
                const result = await response.json();
                showModal(result.title, result.message);
                fetchAndRenderData();
            } else if (newMonto !== null) {
                showModal('Error', 'Monto inválido. Por favor, ingrese un número.');
            }
        }

        async function deletePorCobrar(persona) {
            if (confirm(`¿Confirmas que ${persona} ya ha pagado?`)) {
                const response = await fetch('/eliminar_por_cobrar', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ persona })
                });
                const result = await response.json();
                showModal(result.title, result.message);
                fetchAndRenderData();
            }
        }

        function populateEgresosDeudas(data) {
            const personaSelect = document.getElementById('persona-egreso');
            personaSelect.innerHTML = '<option value="Nadie">Nadie</option>';
            for (const persona in data.deudas) {
                const option = document.createElement('option');
                option.value = persona;
                option.textContent = persona;
                personaSelect.appendChild(option);
            }
        }

        // Manejadores de formulario
        document.getElementById('ingresos-form').addEventListener('submit', async function(e) {
            e.preventDefault();
            const formData = new FormData(e.target);
            const data = Object.fromEntries(formData.entries());
            data.monto = parseFloat(data.monto);

            const response = await fetch('/anadir_ingreso', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });
            const result = await response.json();
            showModal(result.title, result.message);
            showSection('panel');
            e.target.reset();
        });

        document.getElementById('egresos-form').addEventListener('submit', async function(e) {
            e.preventDefault();
            const formData = new FormData(e.target);
            const data = Object.fromEntries(formData.entries());
            data.monto = parseFloat(data.monto);

            const response = await fetch('/anadir_egreso', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });
            const result = await response.json();
            showModal(result.title, result.message);
            showSection('panel');
            e.target.reset();
        });

        document.getElementById('ahorros-form').addEventListener('submit', async function(e) {
            e.preventDefault();
            const formData = new FormData(e.target);
            const data = Object.fromEntries(formData.entries());
            data.monto = parseFloat(data.monto);

            const response = await fetch('/anadir_ahorro', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });
            const result = await response.json();
            showModal(result.title, result.message);
            showSection('panel');
            e.target.reset();
        });

        document.getElementById('deudas-form').addEventListener('submit', async function(e) {
            e.preventDefault();
            const formData = new FormData(e.target);
            const data = Object.fromEntries(formData.entries());
            data.monto = parseFloat(data.monto);

            const response = await fetch('/anadir_deuda', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });
            const result = await response.json();
            showModal(result.title, result.message);
            showSection('deudas');
            e.target.reset();
        });

        function toggleDeudorInput() {
            const metodo = document.getElementById('metodo-ingreso').value;
            const deudorGroup = document.getElementById('deudor-group');
            deudorGroup.style.display = (metodo === 'Crédito') ? 'block' : 'none';
        }

        // Funciones del Modal
        const modalContainer = document.getElementById('modal-container');
        const modalTitle = document.getElementById('modal-title');
        const modalMessage = document.getElementById('modal-message');

        function showModal(title, message) {
            modalTitle.textContent = title;
            modalMessage.textContent = message;
            modalContainer.classList.remove('opacity-0', 'pointer-events-none');
            modalContainer.querySelector('div').classList.remove('scale-95');
            modalContainer.querySelector('div').classList.add('scale-100');
        }

        function closeModal() {
            modalContainer.classList.add('opacity-0', 'pointer-events-none');
            modalContainer.querySelector('div').classList.add('scale-95');
            modalContainer.querySelector('div').classList.remove('scale-100');
        }

        // Carga inicial de datos
        document.addEventListener('DOMContentLoaded', fetchAndRenderData);
    </script>
</body>
</html>
"""

def get_data_from_session():
    """Inicializa la estructura de datos en la sesión si no existe."""
    session.setdefault('ingresos', [])
    session.setdefault('egresos', [])
    session.setdefault('ahorros', [])
    session.setdefault('deudas', {})
    session.setdefault('por_cobrar', {})
    session.setdefault('capital', 0)
    return {
        'ingresos': session['ingresos'],
        'egresos': session['egresos'],
        'ahorros': session['ahorros'],
        'deudas': session['deudas'],
        'por_cobrar': session['por_cobrar'],
        'capital': session['capital']
    }

@app.route('/')
def home():
    """Ruta principal que sirve la aplicación web."""
    return render_template_string(HTML_TEMPLATE)

@app.route('/get_data')
def get_data():
    """Ruta que devuelve todos los datos en formato JSON."""
    return jsonify(get_data_from_session())

@app.route('/anadir_ingreso', methods=['POST'])
def add_income():
    """Ruta para añadir un nuevo ingreso."""
    data = request.json
    monto = data.get('monto')
    procedencia = data.get('procedencia')
    metodo = data.get('metodo')
    deudor = data.get('deudor')

    if not all([monto, procedencia, metodo]):
        return jsonify({"title": "Error", "message": "Faltan datos."}), 400

    new_income = {"monto": monto, "procedencia": procedencia, "metodo": metodo, "deudor": deudor}
    session['ingresos'].append(new_income)
    
    if metodo != "Crédito":
        session['capital'] += monto
    else:
        current_por_cobrar = session.get('por_cobrar', {})
        current_por_cobrar[deudor] = current_por_cobrar.get(deudor, 0) + monto
        session['por_cobrar'] = current_por_cobrar

    return jsonify({"title": "Éxito", "message": "Ingreso registrado."})

@app.route('/anadir_egreso', methods=['POST'])
def add_expense():
    """Ruta para añadir un nuevo egreso."""
    data = request.json
    monto = data.get('monto')
    concepto = data.get('concepto')
    metodo = data.get('metodo')
    persona = data.get('persona')

    if not all([monto, concepto, metodo]):
        return jsonify({"title": "Error", "message": "Faltan datos."}), 400

    new_expense = {"monto": monto, "concepto": concepto, "metodo": metodo, "persona": persona if persona != "Nadie" else None}
    session['egresos'].append(new_expense)
    session['capital'] -= monto
    
    current_deudas = session.get('deudas', {})
    if persona and persona in current_deudas:
        current_deudas[persona] -= monto
        if current_deudas[persona] <= 0:
            del current_deudas[persona]
    
    if metodo == "Crédito" and persona == "Nadie":
        current_deudas[concepto] = current_deudas.get(concepto, 0) + monto

    session['deudas'] = current_deudas

    return jsonify({"title": "Éxito", "message": "Egreso registrado."})

@app.route('/anadir_ahorro', methods=['POST'])
def add_saving():
    """Ruta para añadir un ahorro."""
    data = request.json
    monto = data.get('monto')
    if not monto:
        return jsonify({"title": "Error", "message": "Monto inválido."}), 400

    session['ahorros'].append(monto)
    session['capital'] -= monto
    
    return jsonify({"title": "Éxito", "message": "Ahorro registrado."})

@app.route('/anadir_deuda', methods=['POST'])
def add_debt():
    """Ruta para añadir una nueva deuda."""
    data = request.json
    persona = data.get('persona')
    monto = data.get('monto')

    if not all([persona, monto]):
        return jsonify({"title": "Error", "message": "Faltan datos."}), 400

    current_deudas = session.get('deudas', {})
    current_deudas[persona] = current_deudas.get(persona, 0) + monto
    session['deudas'] = current_deudas
    
    return jsonify({"title": "Éxito", "message": f"Deuda con {persona} añadida por ${monto}."})

@app.route('/editar_por_cobrar', methods=['POST'])
def edit_receivable():
    """Ruta para editar un monto de una cuenta por cobrar."""
    data = request.json
    persona = data.get('persona')
    monto = data.get('monto')
    
    current_por_cobrar = session.get('por_cobrar', {})
    if persona in current_por_cobrar:
        current_por_cobrar[persona] = monto
        session['por_cobrar'] = current_por_cobrar
        return jsonify({"title": "Éxito", "message": f"Monto actualizado para {persona}."})
    else:
        return jsonify({"title": "Error", "message": "Persona no encontrada."}), 404

@app.route('/eliminar_por_cobrar', methods=['POST'])
def delete_receivable():
    """Ruta para eliminar una cuenta por cobrar y actualizar el capital."""
    data = request.json
    persona = data.get('persona')

    current_por_cobrar = session.get('por_cobrar', {})
    if persona in current_por_cobrar:
        session['capital'] += current_por_cobrar[persona]
        del current_por_cobrar[persona]
        session['por_cobrar'] = current_por_cobrar
        return jsonify({"title": "Éxito", "message": f"Deuda con {persona} eliminada y capital actualizado."})
    else:
        return jsonify({"title": "Error", "message": "Persona no encontrada."}), 404

@app.route('/exportar')
def export_to_excel():
    """Ruta para exportar todos los datos a un archivo Excel."""
    data = get_data_from_session()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Finanzas"

    ws.append(["Tipo", "Monto", "Detalle", "Método", "Persona", "Fecha"])

    for i in data['ingresos']:
        ws.append(["Ingreso", i["monto"], i["procedencia"], i["metodo"], i["deudor"], datetime.date.today()])
    for e in data['egresos']:
        ws.append(["Egreso", e["monto"], e["concepto"], e["metodo"], e["persona"], datetime.date.today()])
    for a in data['ahorros']:
        ws.append(["Ahorro", a, "Ahorro personal", "-", "-", datetime.date.today()])
    for persona, monto in data['deudas'].items():
        ws.append(["Deuda (yo debo)", monto, "-", "-", persona, datetime.date.today()])
    for persona, monto in data['por_cobrar'].items():
        ws.append(["Por cobrar (me deben)", monto, "-", "-", persona, datetime.date.today()])

    # Guardar en memoria
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name='finanzas.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True)
