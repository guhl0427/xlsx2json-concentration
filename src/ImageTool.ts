import fs from "fs";
import path from "path";
import child_process from "child_process";

export function handleImage(dir: string) {
    let files = fs.readdirSync(dir);
    for(let file of files){
        if(file === "bg.png") {
            child_process.execSync(`convert bg.png bg.jpg`, {cwd: dir});
        }
    
        /* if(file.indexOf("analyzing") >= 0 && path.extname(file) === ".png") {
            let cmd = `convert ${file} -crop 788x444+0+0 ${file}`;
            child_process.execSync(cmd, {cwd: dir});
            child_process.execSync(`convert -size 788x444 xc:none -draw "roundrectangle 0,0,788,444,32,32" png:- | convert ${file} -matte - -compose DstIn -composite ${file}`, {cwd: dir});
        } */
    }
}