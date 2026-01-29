import * as https from 'https';

// 使用 NHTSA 免費 API 解碼 VIN
async function decodeVin(vin: string): Promise<any> {
  return new Promise((resolve, reject) => {
    const url = `https://vpic.nhtsa.dot.gov/api/vehicles/decodevin/${vin}?format=json`;

    https.get(url, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          resolve(json);
        } catch (e) {
          reject(e);
        }
      });
    }).on('error', reject);
  });
}

// 從 NHTSA 結果中提取關鍵資訊
function extractVehicleInfo(results: any[]) {
  const getValue = (variableId: number) => {
    const item = results.find((r: any) => r.VariableId === variableId);
    return item?.Value || null;
  };

  return {
    make: getValue(26),           // Make
    model: getValue(28),          // Model
    year: getValue(29),           // Model Year
    engineCylinders: getValue(9), // Engine Number of Cylinders
    engineDisplacement: getValue(11), // Displacement (L)
    engineModel: getValue(18),    // Engine Model
    fuelType: getValue(24),       // Fuel Type - Primary
    horsepower: getValue(71),     // Engine Brake (hp) From
    driveType: getValue(15),      // Drive Type
    bodyClass: getValue(5),       // Body Class
    doors: getValue(14),          // Doors
    transmissionStyle: getValue(37), // Transmission Style
  };
}

async function main() {
  // 測試幾個 VIN
  const testVins = [
    'SBM15ACA0KW800014', // McLaren Senna
    'SCBCF13S1NC002895', // Bentley GT V8
    'ZHWEF4ZF8NLA19267', // Lamborghini EVO
    'SCA1S68087UX01093', // RR Phantom
  ];

  for (const vin of testVins) {
    console.log(`\n=== VIN: ${vin} ===`);
    try {
      const result = await decodeVin(vin);
      const info = extractVehicleInfo(result.Results);
      console.log('品牌:', info.make);
      console.log('型號:', info.model);
      console.log('年份:', info.year);
      console.log('引擎汽缸數:', info.engineCylinders);
      console.log('排氣量:', info.engineDisplacement, 'L');
      console.log('引擎型號:', info.engineModel);
      console.log('燃料類型:', info.fuelType);
      console.log('馬力:', info.horsepower, 'hp');
      console.log('驅動類型:', info.driveType);
      console.log('車身類型:', info.bodyClass);
    } catch (e) {
      console.log('解碼失敗:', e);
    }
  }
}

main();
