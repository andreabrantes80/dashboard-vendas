document.addEventListener("DOMContentLoaded", function () {
  let dadosVendedores = [];

  fetch("Vendedor.xlsx")
    .then((response) => response.arrayBuffer())
    .then((data) => {
      const workbook = XLSX.read(data, { type: "array" });
      const nomeDaAba = workbook.SheetNames[0];
      const aba = workbook.Sheets[nomeDaAba];
      dadosVendedores = XLSX.utils.sheet_to_json(aba);
      const vendas = {};
      dadosVendedores.forEach((item) => {
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
          }">

            <div class="podio">

            <div class="posicao position-${index + 1}">

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

      const conteinerRatingVendedores = document.querySelector(
        "#ranking-vendedores"
      );

      conteinerRatingVendedores.innerHTML = vendasOrdenadas
        .map(
          (vendedor, index) => `
            <div class="item-ranking" data-vendedor="${vendedor.vendedor}">
            <img src="images/${vendedor.vendedor}.jpg" alt="${vendedor.vendedor}">
            <div class="info">
                <p>${index + 1} - ${vendedor.vendedor}</p>
                 <p> ${vendedor.total.toLocaleString("pt-BR", {
                   style: "currency",
                   currency: "BRL",
                 })}</p>
            </div>
        </div>
        `
        )
        .join("");

      const containerResumoVendas = document.querySelector(
        "#tabela-resumo tbody"
      );

      containerResumoVendas.innerHTML = dadosVendedores
        .map(
          (venda) => `

            <tr>

            <td>${venda.Vendedor}</td>
            <td>${venda.Produto}</td>
            <td>${venda.Total.toLocaleString("pt-BR", {
              style: "currency",
              currency: "BRL",
            })}</td>

            </tr>

            `
        )
        .join("");

      console.log(dadosVendedores);
    })
    .catch((error) => console.error("Error loading Excel file:", error));

  document
    .getElementById("mostrar-top-3")
    .addEventListener("click", function () {
      document.getElementById("titulo-principal").innerHTML =
        "Melhores Vendedores do mês";
      document.getElementById("top-3-vendedores").style.display = "flex";
      document.getElementById("ranking-vendedores").style.display = "none";
      document.getElementById("resumo-vendas").style.display = "none";
      document.getElementById("detalhes-vendedor").style.display = "none";
    });

  document
    .getElementById("mostrar-ranking")
    .addEventListener("click", function () {
      document.getElementById("titulo-principal").innerHTML =
        "Rankind de Vendas";
      document.getElementById("top-3-vendedores").style.display = "none";
      document.getElementById("ranking-vendedores").style.display = "flex";
      document.getElementById("resumo-vendas").style.display = "none";
      document.getElementById("detalhes-vendedor").style.display = "none";
    });

  document
    .getElementById("mostrar-resumo")
    .addEventListener("click", function () {
      document.getElementById("titulo-principal").innerHTML =
        "Resumo de Vendas";
      document.getElementById("top-3-vendedores").style.display = "none";
      document.getElementById("ranking-vendedores").style.display = "none";
      document.getElementById("resumo-vendas").style.display = "block";
      document.getElementById("detalhes-vendedor").style.display = "none";
    });
});
