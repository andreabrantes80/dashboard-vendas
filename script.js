document.addEventListener('DOMContentLoaded', function () {
    let dadosVendedores = []

    fetch('Vendedor.xlsx')
    .then(response => response.arrayBuffer() )
    .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const nomeDaAba = workbook.SheetNames[0];
        const aba = workbook.Sheets[nomeDaAba];
        dadosVendedores = XLSX.utils.sheet_to_json(aba);
        const vendas = {};
        dadosVendedores.forEach(item => {
            const vendedor = item.Vendedor;
            const total = parseFloat(item.Total);
            if (vendas[vendedor]) {
                vendas[vendedor] += total;
            } else {
                vendas[vendedor] = total;
            }
        });
        const vendasOrdenadas = Object.entries(vendas)
            .map(([vendedor, total]) => ({ vendedor, total }))
        .sort((a, b) => b.total - a.total);
        const topVendedores = vendasOrdenadas.slice(0, 3);
        const containerVendedores = document.querySelector("#top-3-vendedores");
        containerVendedores.innerHTML = topVendedores
          .map(
            (vendedor, index) => `
            <div class="vendedor">

            <img src="images/${vendedor.vendedor}.jpg" alt="${
              vendedor.vendedor
            }" class="foto-vendedor">

            <div class="podio">

            <div class="posicao posicao-${index + 1}">

            <p>${index + 1}</p>

            </div>

            <h3>${vendedor.vendedor}</h3>

            <p> ${vendedor.total.toLocaleString("pt-BR", {
              style: "currency",
              currency: "BRL",
            })}</p>

            </div>

            </div>

        `
          )
          .join("");

        const conatinerRatingVendedores = document.querySelector('#ranking-vendedores');

        conatinerRatingVendedores.innerHTML = vendasOrdenadas.map((vendedor, index) => `

            <div class="item-ranking" data-vendedor="${vendedor.vendedor}">

            <img src="images/${vendedor.vendedor}.jpg" alt="${vendedor.vendedor}" >

            <div class="info">

            <p>${index + 1} - ${vendedor.vendedor}</p>

            <p> ${vendedor.total.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'})}</p>

            </div>

            </div

        `).join('');

        const containerResumoVendas = document.querySelector('tabela-resumo');

        containerResumoVendas.innerHTML = dadosVendedores.map(venda => `

            <tr>

            <td>${venda.vendedor}</td>
            <td>${venda.produto}</td>
            <td>${venda.Total.toLocaleString('pt-BR', {style:'currency', currency: 'BRL'})}</td>

            </tr>

            `).join('');

        console.log(dadosVendedores);
    })
        .catch(error => console.error('Error loading Excel file:', error));


});