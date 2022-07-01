import minimist from "minimist";
import path from "path";
import { parseFile } from "./ConfigTool";
import { handleImage } from "./ImageTool";

const assetDir = `D:\\work\\excel`;

const params = minimist(process.argv.slice(2));
if(params.id === undefined) {
    console.error("没有输入参数课程id !!!");
}

parseFile(path.join(assetDir, `lesson${params.id}.xlsx`));
// handleImage(path.join(assetDir, `lesson${params.id}`));