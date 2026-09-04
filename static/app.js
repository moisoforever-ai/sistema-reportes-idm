document.addEventListener('DOMContentLoaded', () => {
    const loaderInitial = document.getElementById('loader-initial');
    const loaderError = document.getElementById('loader-error');
    const retryOptionsBtn = document.getElementById('retry-options-btn');
    const reportForm = document.getElementById('report-form');
    const selectMes = document.getElementById('mes');
    const selectFecha = document.getElementById('fecha');
    const selectEmpresa = document.getElementById('empresa');
    const selectSucursal = document.getElementById('sucursal');
    const inputAperturaEl = document.getElementById('hora_apertura');
    const inputCierreEl = document.getElementById('hora_cierre');
    const horarioError = document.getElementById('horario-error');
    const btnSubmit = document.getElementById('submit-btn');
    const btnText = btnSubmit.querySelector('.btn-text');
    const btnLoader = btnSubmit.querySelector('.btn-loader');
    const btnObs = document.getElementById('obs-btn');
    const btnObsText = btnObs.querySelector('.btn-text');
    const btnObsLoader = btnObs.querySelector('.btn-loader');
    const toast = document.getElementById('toast');
    const toastMessage = document.getElementById('toast-message');

    let apiData = null;

    const mesNombres = {
        '01': 'Enero', '02': 'Febrero', '03': 'Marzo', '04': 'Abril',
        '05': 'Mayo', '06': 'Junio', '07': 'Julio', '08': 'Agosto',
        '09': 'Septiembre', '10': 'Octubre', '11': 'Noviembre', '12': 'Diciembre'
    };

    // Toast utility
    function showToast(message, type = 'success') {
        toastMessage.textContent = message;
        toast.className = `toast show ${type}`;
        
        setTimeout(() => {
            toast.className = 'toast hidden';
        }, 4000);
    }

    // Fix UX: valida que la hora de cierre sea posterior a la de apertura. Antes no
    // había ningún aviso; el reporte se generaba igual pero con la proyección en 0%,
    // de forma silenciosa. Se llama en cada cambio de horario y también antes de
    // enviar el formulario (devuelve true/false para poder bloquear el envío).
    function validarHorarios() {
        const apertura = inputAperturaEl.value;
        const cierre = inputCierreEl.value;
        const valido = !apertura || !cierre || cierre > apertura;
        horarioError.classList.toggle('hidden', valido);
        return valido;
    }
    inputAperturaEl.addEventListener('change', validarHorarios);
    inputCierreEl.addEventListener('change', validarHorarios);

    // Fix UX: antes, si esta llamada fallaba (Google Sheets caído, red lenta), el
    // spinner "Cargando parámetros..." quedaba girando para siempre — el único aviso
    // era un toast que desaparecía solo en 4 segundos, sin forma de reintentar sin
    // recargar toda la página. Ahora se puede reintentar con un botón.
    function cargarOpciones() {
        loaderError.classList.add('hidden');
        loaderInitial.classList.remove('hidden');

        fetch('/api/options')
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    apiData = data;
                    
                    // Extract unique months from dates
                    const uniqueMonths = new Set();
                    data.dates.forEach(date => {
                        const parts = date.split('/');
                        if (parts.length === 3) {
                            uniqueMonths.add(parts[1]); // MM part
                        }
                    });

                    // Sort unique months chronologically
                    const sortedMonths = Array.from(uniqueMonths).sort((a, b) => parseInt(a) - parseInt(b));

                    // Populate Meses
                    selectMes.innerHTML = '<option value="" disabled selected>Seleccione un mes</option>';
                    sortedMonths.forEach(m => {
                        const option = document.createElement('option');
                        option.value = m;
                        option.textContent = mesNombres[m] || `Mes ${m}`;
                        selectMes.appendChild(option);
                    });

                    // Populate Empresas
                    selectEmpresa.innerHTML = '<option value="" disabled selected>Seleccione una empresa</option>';
                    data.companies.forEach(company => {
                        const option = document.createElement('option');
                        option.value = company;
                        option.textContent = company;
                        selectEmpresa.appendChild(option);
                    });

                    // Transition UI
                    loaderInitial.classList.add('hidden');
                    reportForm.classList.remove('hidden');
                } else {
                    loaderInitial.classList.add('hidden');
                    loaderError.classList.remove('hidden');
                    showToast(`Error: ${data.error}`, 'error');
                }
            })
            .catch(err => {
                loaderInitial.classList.add('hidden');
                loaderError.classList.remove('hidden');
                showToast('Error de conexión al cargar datos de Google Sheets.', 'error');
                console.error(err);
            });
    }

    retryOptionsBtn.addEventListener('click', cargarOpciones);

    // Fetch parameters from Google Sheets via backend
    cargarOpciones();

    // Reactive filter of Dates based on Selected Month
    selectMes.addEventListener('change', () => {
        const selectedMes = selectMes.value;
        selectFecha.innerHTML = '<option value="" disabled selected>Seleccione el día</option>';
        selectFecha.disabled = true;
        
        if (selectedMes && apiData && apiData.dates) {
            // Filter dates belonging to this month
            const filteredDates = apiData.dates.filter(date => {
                const parts = date.split('/');
                return parts[1] === selectedMes;
            });
            
            // Sort dates chronologically
            filteredDates.sort((a, b) => {
                const parseDate = str => {
                    const p = str.split('/');
                    return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
                };
                return parseDate(a) - parseDate(b);
            });
            
            if (filteredDates.length > 0) {
                filteredDates.forEach(date => {
                    const option = document.createElement('option');
                    option.value = date;
                    option.textContent = date;
                    selectFecha.appendChild(option);
                });
                selectFecha.disabled = false;
            } else {
                selectFecha.innerHTML = '<option value="" disabled selected>No hay fechas este mes</option>';
            }
        }
        
        updateBranches();
    });

    // Update branches dropdown reactively based on selected date and company
    function updateBranches() {
        const selectedDate = selectFecha.value;
        const selectedCompany = selectEmpresa.value;
        
        selectSucursal.innerHTML = '<option value="" disabled selected>Seleccione una sucursal</option>';
        selectSucursal.disabled = true;
        
        if (!selectedDate || !selectedCompany) {
            if (!selectedDate && !selectedCompany) {
                selectSucursal.innerHTML = '<option value="" disabled selected>Seleccione fecha y empresa</option>';
            } else if (!selectedDate) {
                selectSucursal.innerHTML = '<option value="" disabled selected>Seleccione una fecha</option>';
            } else {
                selectSucursal.innerHTML = '<option value="" disabled selected>Seleccione una empresa</option>';
            }
            return;
        }

        if (apiData && apiData.branches[selectedDate] && apiData.branches[selectedDate][selectedCompany]) {
            const branches = apiData.branches[selectedDate][selectedCompany];
            if (branches && branches.length > 0) {
                branches.forEach(branch => {
                    const option = document.createElement('option');
                    option.value = branch;
                    option.textContent = branch.replace(/_/g, ' ').toUpperCase();
                    selectSucursal.appendChild(option);
                });
                selectSucursal.disabled = false;
            } else {
                selectSucursal.innerHTML = '<option value="" disabled selected>No hay tiendas estudiadas este día</option>';
            }
        } else {
            selectSucursal.innerHTML = '<option value="" disabled selected>No hay tiendas estudiadas este día</option>';
        }
    }

    selectFecha.addEventListener('change', updateBranches);
    selectEmpresa.addEventListener('change', updateBranches);

    // Handle form submit to download Excel file
    reportForm.addEventListener('submit', (e) => {
        e.preventDefault();

        if (!validarHorarios()) {
            showToast('La hora de cierre debe ser posterior a la de apertura.', 'error');
            return;
        }

        // Get values
        const payload = {
            fecha: selectFecha.value,
            empresa: selectEmpresa.value,
            sucursal: selectSucursal.value,
            hora_apertura: document.getElementById('hora_apertura').value,
            hora_cierre: document.getElementById('hora_cierre').value
        };

        // UI loading state
        btnSubmit.disabled = true;
        btnText.classList.add('hidden');
        btnLoader.classList.remove('hidden');

        fetch('/api/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        })
        .then(async response => {
            if (response.ok) {
                // Get filename from header
                const disposition = response.headers.get('content-disposition');
                let filename = 'Reporte_Auditoria.xlsx';
                if (disposition && disposition.indexOf('attachment') !== -1) {
                    const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    const matches = filenameRegex.exec(disposition);
                    if (matches != null && matches[1]) { 
                        filename = matches[1].replace(/['"]/g, '');
                    }
                }

                // Download blob
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
                showToast('¡Reporte generado e iniciando descarga con éxito!', 'success');
            } else {
                let errMsg = 'Ocurrió un error al procesar el reporte.';
                try {
                    const errData = await response.json();
                    errMsg = errData.error || errMsg;
                } catch(e) {
                    try {
                        const text = await response.text();
                        if (text && text.length < 200) {
                            errMsg = text;
                        } else {
                            errMsg = `Error del servidor (Estado ${response.status}): ${response.statusText}`;
                        }
                    } catch(e2) {
                        errMsg = `Error del servidor (Estado ${response.status})`;
                    }
                }
                showToast(errMsg, 'error');
            }
        })
        .catch(err => {
            showToast('Error de comunicación: ' + err.message, 'error');
            console.error(err);
        })
        .finally(() => {
            // Restore UI state
            btnSubmit.disabled = false;
            btnText.classList.remove('hidden');
            btnLoader.classList.add('hidden');
        });
    });

    // Handle observations button click
    btnObs.addEventListener('click', () => {
        if (!reportForm.checkValidity()) {
            reportForm.reportValidity();
            return;
        }
        if (!validarHorarios()) {
            showToast('La hora de cierre debe ser posterior a la de apertura.', 'error');
            return;
        }

        const payload = {
            fecha: selectFecha.value,
            empresa: selectEmpresa.value,
            sucursal: selectSucursal.value,
            hora_apertura: document.getElementById('hora_apertura').value,
            hora_cierre: document.getElementById('hora_cierre').value
        };

        // UI loading state
        btnObs.disabled = true;
        btnObsText.classList.add('hidden');
        btnObsLoader.classList.remove('hidden');

        fetch('/api/generate_observations', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        })
        .then(async response => {
            if (response.ok) {
                // Get filename from header
                const disposition = response.headers.get('content-disposition');
                let filename = 'Observaciones_Auditoria.txt';
                if (disposition && disposition.indexOf('attachment') !== -1) {
                    const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    const matches = filenameRegex.exec(disposition);
                    if (matches != null && matches[1]) { 
                        filename = matches[1].replace(/['"]/g, '');
                    }
                }

                // Download blob
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
                showToast('¡Observaciones generadas e iniciando descarga!', 'success');
            } else {
                let errMsg = 'Ocurrió un error al generar las observaciones.';
                try {
                    const errData = await response.json();
                    errMsg = errData.error || errMsg;
                } catch(e) {
                    try {
                        const text = await response.text();
                        if (text && text.length < 200) {
                            errMsg = text;
                        } else {
                            errMsg = `Error del servidor (Estado ${response.status}): ${response.statusText}`;
                        }
                    } catch(e2) {
                        errMsg = `Error del servidor (Estado ${response.status})`;
                    }
                }
                showToast(errMsg, 'error');
            }
        })
        .catch(err => {
            showToast('Error de comunicación: ' + err.message, 'error');
            console.error(err);
        })
        .finally(() => {
            // Restore UI state
            btnObs.disabled = false;
            btnObsText.classList.remove('hidden');
            btnObsLoader.classList.add('hidden');
        });
    });

    // Handle database sync button click
    const btnSync = document.getElementById('sync-btn');
    if (btnSync) {
        btnSync.addEventListener('click', () => {
            btnSync.disabled = true;
            btnSync.classList.add('spinning');
            const syncText = btnSync.querySelector('.sync-text');
            const originalText = syncText.textContent;
            syncText.textContent = 'Sincronizando...';

            fetch('/api/sync', {
                method: 'POST'
            })
            .then(async response => {
                if (response.ok) {
                    showToast('¡Base de datos sincronizada con Google Sheets con éxito!', 'success');
                    // Reload page or re-load option data to reflect changes
                    window.location.reload();
                } else {
                    let errMsg = 'No se pudo conectar con Google Sheets.';
                    try {
                        const errData = await response.json();
                        errMsg = errData.error || errMsg;
                    } catch(e) {
                        try {
                            const text = await response.text();
                            if (text && text.length < 200) {
                                errMsg = text;
                            } else {
                                errMsg = `Error del servidor (Estado ${response.status}): ${response.statusText}`;
                            }
                        } catch(e2) {
                            errMsg = `Error del servidor (Estado ${response.status})`;
                        }
                    }
                    showToast(`${errMsg}\nSe continuará usando la base de datos local cached.`, 'warning');
                }
            })
            .catch(err => {
                showToast(`Error de red: ${err.message}\nSe continuará usando la base de datos local cached.`, 'warning');
            })
            .finally(() => {
                btnSync.disabled = false;
                btnSync.classList.remove('spinning');
                syncText.textContent = originalText;
            });
        });
    }
});
