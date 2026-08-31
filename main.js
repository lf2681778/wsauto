const fs = require('fs');
const XLSX = require('xlsx');

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const CHAIN_ID = process.env.CHAIN_ID;

// === 你可以在這裡修改容忍值 ===
const ERROR_THRESHOLD = 50; 
let errorCount = 0;
let successCount = 0; 

async function getToken() {
    if (!CLIENT_ID || !CLIENT_SECRET) {
        console.log("❌ 嚴重錯誤：GitHub Secrets 沒有被正確讀取！");
        process.exit(1); 
    }

    const params = new URLSearchParams();
    params.append('grant_type', 'client_credentials');
    params.append('client_id', CLIENT_ID);
    params.append('client_secret', CLIENT_SECRET);

    try {
        const res = await fetch('https://foodpanda.partner.deliveryhero.io/v2/oauth/token', {
            method: 'POST', body: params
        });
        const data = await res.json();
        if (data.access_token) return data.access_token;
        
        console.log("❌ 取得 Token 失敗，Foodpanda 回傳：", data);
        process.exit(1); 
    } catch (err) {
        console.log("❌ 請求 Token 時發生網路錯誤：", err.message);
        process.exit(1);
    }
}

async function updateVendor(token, vendorId) {
    const url = `https://foodpanda.partner.deliveryhero.io/v2/chains/${CHAIN_ID}/vendors/${vendorId}/status`;
    try {
        const res = await fetch(url, {
            method: 'PUT',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ status: "OPEN" })
        });
        
        if (res.ok) {
            console.log(`✅ [${vendorId}] 更新成功`);
            successCount++; 
        } else {
            console.log(`❌ [${vendorId}] 更新失敗: ${await res.text()}`);
            errorCount++; 
        }
    } catch (err) {
        console.log(`❌ [${vendorId}] 網路錯誤: ${err.message}`);
        errorCount++; 
    }
}

async function main() {
    const token = await getToken();
    const workbook = XLSX.readFile('VendorWatsons.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    let vendors = [];
    for (const row of rawData) {
        const id = row[0] ? row[0].toString().trim() : '';
        if (id && id.toLowerCase() !== 'vendor_id') vendors.push(id);
    }

    console.log(`共讀取到 ${vendors.length} 家店家，準備開始執行...`);

    for (let i = 0; i < vendors.length; i++) {
        const vid = vendors[i];
        await updateVendor(token, vid);
        await new Promise(r => setTimeout(r, 500)); 
    }
    
    // === 將結果數字寫入，交接給 GitHub Actions ===
    if (process.env.GITHUB_OUTPUT) {
        fs.appendFileSync(process.env.GITHUB_OUTPUT, `success_count=${successCount}\n`);
        fs.appendFileSync(process.env.GITHUB_OUTPUT, `error_count=${errorCount}\n`);
    }

    // === 最終檢查邏輯 ===
    if (errorCount > ERROR_THRESHOLD) {
        console.log(`\n🚨 警告：共有 ${errorCount} 家店鋪更新失敗，已超過容忍值！`);
        process.exit(1); 
    } else {
        console.log(`\n🎉 執行完畢！成功: ${successCount}, 失敗: ${errorCount}`);
        process.exit(0);
    }
}

main();
