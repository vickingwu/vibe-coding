var https = require('https');

function get(url, headers) {
    return new Promise(function(resolve, reject) {
        var opts = typeof url === 'string' ? url : url;
        if (headers) {
            var u = new URL(url);
            opts = { hostname: u.hostname, path: u.pathname + u.search, headers: headers };
        }
        https.get(opts, function(res) {
            var d = '';
            res.on('data', function(c) { d += c; });
            res.on('end', function() { resolve({ status: res.statusCode, headers: res.headers, body: d }); });
        }).on('error', reject);
    });
}

async function main() {
    // Test 1: 集思录 LOF API - 看看申购状态字段
    console.log('=== 集思录 LOF API ===');
    try {
        var r = await get('https://www.jisilu.cn/data/lof/stock_lof_list/?___jsl=LST___t=' + Date.now(), {
            'Referer': 'https://www.jisilu.cn/data/lof/',
            'User-Agent': 'Mozilla/5.0'
        });
        var json = JSON.parse(r.body);
        if (json.rows && json.rows.length > 0) {
            // 找我们的基金
            var codes = ['161226', '160416', '501018', '161125'];
            var found = json.rows.filter(function(row) { return codes.indexOf(row.id) >= 0; });
            found.forEach(function(row) {
                var c = row.cell;
                console.log(row.id, c.fund_nm, 'apply_fee=' + c.apply_fee, 'apply_status=' + c.apply_status,
                    'discount_rt=' + c.discount_rt, 'price=' + c.price, 'fund_nav=' + c.fund_nav);
            });
            // 也看看第一条的所有字段
            console.log('\n第一条完整字段:');
            var first = json.rows[0].cell;
            Object.keys(first).forEach(function(k) { console.log('  ' + k + '=' + first[k]); });
        }
        console.log('总记录数:', json.rows ? json.rows.length : 0);
    } catch(e) {
        console.log('集思录失败:', e.message);
    }

    // Test 2: fundmobapi 带 callback 看看有没有申购状态
    console.log('\n=== fundmobapi 详情 ===');
    try {
        var r2 = await get('https://fundmobapi.eastmoney.com/FundMNewApi/FundMNFInfo?pageIndex=1&pageSize=5&plat=Android&appType=ttjj&product=EFund&Version=1&deviceid=1&Fcodes=161226,160416,501018');
        var json2 = JSON.parse(r2.body);
        json2.Datas.forEach(function(d) {
            console.log(d.FCODE, d.SHORTNAME, 'NAV=' + d.NAV);
            // 打印所有字段
            Object.keys(d).forEach(function(k) { if (d[k] != null && d[k] !== '' && d[k] !== false) console.log('  ' + k + '=' + d[k]); });
        });
    } catch(e) {
        console.log('fundmobapi失败:', e.message);
    }

    // Test 3: 天天基金 lsjz API (有 SGZT 字段)
    console.log('\n=== 天天基金 lsjz API ===');
    try {
        var r3 = await get('https://api.fund.eastmoney.com/f10/lsjz?fundCode=161226&pageIndex=1&pageSize=1', {
            'Referer': 'https://fund.eastmoney.com/'
        });
        var json3 = JSON.parse(r3.body);
        console.log('CORS headers:', r3.headers['access-control-allow-origin'] || 'none');
        if (json3.Data && json3.Data.LSJZList) {
            json3.Data.LSJZList.forEach(function(item) {
                console.log('SGZT=' + item.SGZT, 'SHZT=' + item.SHZT, 'DWJZ=' + item.DWJZ, 'FSRQ=' + item.FSRQ);
            });
        }
    } catch(e) {
        console.log('lsjz失败:', e.message);
    }
}

main();
