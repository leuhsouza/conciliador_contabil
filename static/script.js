document.addEventListener('DOMContentLoaded', function() {
    // Inicializa a soma total com base nos valores conciliados na tela
    let totalSum = 0;
    let selectedIds = []; // Lista para armazenar os IDs das linhas destacadas

    // Calcula a soma inicial com base nas linhas conciliadas (classe 'highlight')
    document.querySelectorAll('table.data tr.highlight').forEach(function(row) {
        const dValue = parseFloat(row.querySelector('td:nth-child(7)').textContent.replace(',', '.')) || 0; // Coluna D
        const cValue = parseFloat(row.querySelector('td:nth-child(8)').textContent.replace(',', '.')) || 0; // Coluna C
        totalSum += dValue - cValue;
    });
    document.getElementById('sumValue').textContent = totalSum.toFixed(2);

    function updateSum(row, isAdding) {
        const dValue = parseFloat(row.querySelector('td:nth-child(7)').textContent.replace(',', '.')) || 0; // Coluna D
        const cValue = parseFloat(row.querySelector('td:nth-child(8)').textContent.replace(',', '.')) || 0; // Coluna C

        if (isAdding) {
            totalSum += dValue;
            totalSum -= cValue;
        } else {
            totalSum -= dValue;
            totalSum += cValue;
        }

        document.getElementById('sumValue').textContent = totalSum.toFixed(2);
    }

    let selectedRow = null;

    document.querySelectorAll('table.data tr').forEach(function(row) {
        row.addEventListener('click', function() {
            if (selectedRow) {
                selectedRow.classList.remove('selected');
            }
            selectedRow = this;
            this.classList.add('selected');
            this.focus();
        });

        row.addEventListener('dblclick', function() {
            const rowId = this.querySelector('td:first-child').textContent; // Supondo que o ID esteja na primeira coluna
            if (this.classList.contains('highlight')) {
                updateSum(this, false);
                this.classList.remove('highlight');
                selectedIds = selectedIds.filter(id => id !== rowId); // Remove o ID da lista
            } else {
                updateSum(this, true);
                this.classList.add('highlight');
                selectedIds.push(rowId); // Adiciona o ID à lista
            }

            // Atualiza o campo oculto com os IDs das linhas destacadas
            const selectedIdsInput = document.getElementById('selectedIds');
            if (selectedIdsInput) {
                selectedIdsInput.value = selectedIds.join(',');
            }
        });
    });

    document.addEventListener('keydown', function(event) {
        if (!selectedRow) return;

        const rows = Array.from(document.querySelectorAll('table.data tr'));
        const index = rows.indexOf(selectedRow);

        if (event.key === 'ArrowDown') {
            event.preventDefault();
            if (index < rows.length - 1) {
                selectedRow.classList.remove('selected');
                selectedRow = rows[index + 1];
                selectedRow.classList.add('selected');
                selectedRow.focus();
            }
        } else if (event.key === 'ArrowUp') {
            event.preventDefault();
            if (index > 0) {
                selectedRow.classList.remove('selected');
                selectedRow = rows[index - 1];
                selectedRow.classList.add('selected');
                selectedRow.focus();
            }
        } else if (event.key === 'Enter') {
            event.preventDefault();
            selectedRow.dispatchEvent(new Event('dblclick'));
        }
    });
});

function saveConciliation() {
    let selectedIds = [];
    let allIds = [];

    // Captura os IDs de todas as linhas exibidas na tabela
    document.querySelectorAll('table.data tr').forEach(function(row) {
        const idElement = row.querySelector('td:first-child');
        if (idElement) {
            const id = idElement.textContent;
            if (id) {
                allIds.push(id.trim());
                if (row.classList.contains('highlight')) {
                    selectedIds.push(id.trim());
                }
            }
        }
    });

    if (allIds.length > 0) {
        fetch('/save_conciliation', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ ids: selectedIds, all_ids: allIds })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                alert('Conciliação salva com sucesso!');
                window.location.reload(); // Atualiza a página para mostrar as alterações
            } else {
                alert('Erro ao salvar a conciliação: ' + data.message);
            }
        })
        .catch(() => {
            alert('Erro ao salvar a conciliação.');
        });
    } else {
        alert('Nenhuma linha encontrada para salvar.');
    }
}

// Função para remover a conciliação das linhas filtradas (apenas visual)
function removeConciliation() {
    const filterField = document.getElementById('filter_field').value.trim();
    const filterConta = document.getElementById('filter_conta').value.trim();
    let rowsUpdated = false; // Verifica se alguma linha foi atualizada

    document.querySelectorAll('table.data tr').forEach(row => {
        const historico = row.querySelector('td:nth-child(3)')?.textContent || '';
        const conta = row.querySelector('td:nth-child(10)')?.textContent || ''; // Ajustar conforme o índice da coluna

        if ((filterField === '' || historico.includes(filterField)) &&
            (filterConta === '' || conta.includes(filterConta))) {
            if (row.classList.contains('highlight')) {
                row.classList.remove('highlight');
                rowsUpdated = true; // Marca que houve atualização
            }
        }
    });

    // Se todas as linhas destacadas forem removidas, zera a soma
    if (rowsUpdated) {
        totalSum = 0; // Zera a soma total
        document.getElementById('sumValue').textContent = totalSum.toFixed(2);
    }

    alert('Conciliação removida visualmente. Para salvar as alterações, clique em "Salvar Conciliação".');
}

function updateTableSum() {
    let totalD = 0;
    let totalC = 0;

    // Itera sobre todas as linhas visíveis na tabela e soma os valores das colunas D e C
    document.querySelectorAll('table.data tbody tr').forEach(row => {
        const dValue = parseFloat(row.querySelector('td:nth-child(7)').textContent.replace(',', '.')) || 0; // Coluna D
        const cValue = parseFloat(row.querySelector('td:nth-child(8)').textContent.replace(',', '.')) || 0; // Coluna C
        
        totalD += dValue;
        totalC += cValue;
    });

    // Atualiza os valores exibidos no rodapé
    document.getElementById('sumColD').textContent = totalD.toFixed(2);
    document.getElementById('sumColC').textContent = totalC.toFixed(2);
}

// Chame a função ao carregar a página e ao aplicar filtros
document.addEventListener('DOMContentLoaded', updateTableSum);
document.querySelector('form').addEventListener('submit', function() {
    // Chame updateTableSum após aplicar o filtro
    setTimeout(updateTableSum, 500); // Delay para esperar a atualização da tabela
});