var lineChartData = {
    labels: ["3", "6", "9", "12", "15", "18", "21", "24", "27"], //顯示區間名稱
    datasets: [{
        label: '未曾使用', // tootip 出現的名稱
        lineTension: 0, // 曲線的彎度，設0 表示直線
        backgroundColor: "#ea464d",
        borderColor: "#ea464d",
        borderWidth: 5,
        data: [10, 12, 15, 18, 22, 33, 50, 60, 130], // 資料
        fill: false, // 是否填滿色彩
    }, {
        label: '罹癌後開始使用',
        lineTension: 0,
        fill: false,
        backgroundColor: "#29b288",
        borderColor: "#29b288",
        borderWidth: 5,
        data: [12, 14, 18, 20, 21, 34, 60, 80, 200],
    },]
};
function drawLineCanvas(ctx,data) {
    window.myLine = new Chart(ctx, {  //先建立一個 chart
        type: 'line', // 型態
        data: data,
        options: {
                responsive: true,
                legend: { //是否要顯示圖示
                    display: true,
                },
                tooltips: { //是否要顯示 tooltip
                    enabled: true
                },
                scales: {  //是否要顯示 x、y 軸
                    xAxes: [{
                        display: true
                    }],
                    yAxes: [{
                        display: true
                    }]
                },
            }
    });
};
window.onload = function () {
    var ctx = document.getElementById("canvas").getContext("2d");
    drawLineCanvas(ctx,lineChartData);
};