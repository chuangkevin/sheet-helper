import * as fs from 'fs';
import * as path from 'path';

export interface CarData {
  name: string;
  folder: string;
  postHelper?: string;
  facebook?: string;
  yahoo?: string;
  official?: string;
  [key: string]: any;
}

export function parseCarPrompts(basePath: string): CarData[] {
  const cars: CarData[] = [];
  const carsPath = path.join(basePath, '汽車資料');

  if (!fs.existsSync(carsPath)) {
    console.warn('Car prompts path not found:', carsPath);
    return cars;
  }

  const folders = fs.readdirSync(carsPath, { withFileTypes: true })
    .filter(dirent => dirent.isDirectory())
    .map(dirent => dirent.name);

  for (const folder of folders) {
    const folderPath = path.join(carsPath, folder);
    const carData: CarData = {
      name: folder,
      folder: folderPath,
    };

    // Read post-helper.md if exists
    const postHelperPath = path.join(folderPath, 'post-helper.md');
    if (fs.existsSync(postHelperPath)) {
      carData.postHelper = fs.readFileSync(postHelperPath, 'utf-8');
    }

    // Read other platform files
    const platforms = ['Facebook.md', 'Yahoo.md', '官方網站.md', '8891.md'];
    for (const platform of platforms) {
      const filePath = path.join(folderPath, platform);
      if (fs.existsSync(filePath)) {
        const key = platform.replace('.md', '').toLowerCase();
        carData[key] = fs.readFileSync(filePath, 'utf-8');
      }
    }

    cars.push(carData);
  }

  return cars;
}