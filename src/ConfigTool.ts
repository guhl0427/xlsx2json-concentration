import fs from "fs";
import path from "path";
import xlsx from "xlsx";

type Prop = {
    name: string,
    offset?: number,
    checkNull?: boolean,
    actions?: Function[]
    complex?: {
        count: number,
        props: Prop[]
    }
};

let props: Prop[] = [
    { name: "id", checkNull: true, actions: [parseInt] },
    { name: "bg" },
    { name: "questionContent" },
    { name: "questionAudio" },
    { name: "picture" },
    {
        name: "items", complex: {
            count: 6,
            props: [
                { name: "index", checkNull: true, actions: [subOne] },
                { name: "sceneUrl" },
                { name: "optionUrl" },
                { name: "scenePos", actions: [splitToNumbers] },
                { name: "isRight", actions: [parseBoolean] }
            ]
        }
    },
    { name: "optionCount", actions: [parseInt] },
    { name: "wrongAudio" },
    { name: "guideAudio" },
    { name: "finishAudio" },
    { name: "effectId" }
];

let props2: Prop[] = [
    { name: "id", checkNull: true, actions: [parseInt] },
    { name: "bg" },
    { name: "questionContent" },
    { name: "questionAudio" },
    {
        name: "items", complex: {
            count: 4,
            props: [
                { name: "index", checkNull: true, actions: [subOne] },
                { name: "audioUrl" },
                { name: "idleUrl" },
                { name: "talkUrl" },
            ]
        }
    },
    { name: "finishAudio" },
    { name: "effectId" }
];

let props3: Prop[] = [
    { name: "id", checkNull: true, actions: [parseInt] },
    { name: "bg" },
    { name: "questionContent" },
    { name: "questionAudio" },
    {
        name: "items", complex: {
            count: 4,
            props: [
                { name: "index", checkNull: true, actions: [subOne] },
                { name: "imgUrl" }
            ]
        }
    },
    { name: "rightItemIndex", actions: [subOne] },
    { name: "rightAudio" },
    { name: "wrongAudio" },
    { name: "finishAudio" },
    { name: "effectId" }
];

let props4: Prop[] = [
    { name: "id", checkNull: true, actions: [parseInt] },
    { name: "bg" },
    { name: "questionContent" },
    { name: "questionAudio" },
    { name: "questionImage" },
    { name: "questionImagePos", actions: [splitToNumbers] },
    { name: "optionPos", actions: [splitToNumbers] },
    {
        name: "items", complex: {
            count: 4,
            props: [
                { name: "index", checkNull: true, actions: [subOne] },
                { name: "optionUrl" },
                { name: "sceneUrl" },
                { name: "scenePos", actions: [splitToNumbers] },
                { name: "isRight", actions: [parseBoolean] },
            ]
        }
    },
    { name: "rightItemIndex", actions: [subOne] },
    { name: "rightImage" },
    { name: "rightAudio" },
    { name: "wrongAudio" },
    { name: "finishAudio" },
    { name: "effectId" }
];

let props5: Prop[] = [
    { name: "id", checkNull: true, actions: [parseInt] },
    { name: "subType", actions: [subOne] },
    { name: "bg" },
    { name: "questionContent" },
    { name: "questionAudio" },
    { name: "questionImage" },
    {
        name: "targets", complex: {
            count: 4,
            props: [
                { name: "index", checkNull: true, actions: [subOne] },
                { name: "imageUrl" },
                { name: "bounds", actions: [splitToNumbers] }
            ]
        }
    },
    {
        name: "items", complex: {
            count: 4,
            props: [
                { name: "index", checkNull: true, actions: [subOne] },
                { name: "imageUrl" },
                { name: "targetId", actions: [subOne] }
            ]
        }
    },
    {
        name: "explain", complex: {
            count: 4,
            props: [
                { name: "index", checkNull: true, actions: [subOne] },
                { name: "imageUrl" },
                { name: "audioUrl" }
            ]
        }
    },
    { name: "rightAudio" },
    { name: "wrongAudio" },
    { name: "allRightAudio" },
    { name: "finishAudio" },
    { name: "effectId" }
];

let props7: Prop[] = [
    { name: "id", checkNull: true, actions: [parseInt] },
    { name: "subType" },
    { name: "bg" },
    { name: "questionAudio" },
    { name: "picture" },
    { name: "picturePos", actions: [splitToNumbers] },
    {
        name: "items", complex: {
            count: 4,
            props: [
                { name: "scenePos", actions: [splitToNumbers] },
                { name: "isRight", actions: [parseBoolean] },
            ]
        }
    },
    { name: "rightTagPos", actions: [splitToNumbers] },
    { name: "completeAudio" },
    { name: "wrongAudio" },
    { name: "finishAudio" },
    { name: "effectId" }
];


let propMap: { [key: number]: Prop[] } = {
    2: props,
    3: props2,
    4: props3,
    5: props4,
    6: props5,
    7: props7
};

function parse(worksheet: xlsx.WorkSheet, props: Prop[]): any {
    let configs: any[] = [];
    out: for (let i = 3; ; i++) {
        let info: any = {};
        let position = 0;
        for (let prop of props) {
            if (prop.complex) {
                info[prop.name] = [];
                position++;
                let count = getWorksheetValue(worksheet, `${numToColumn(position)}${i}`);
                for (let j = 0; j < count; j++) {
                    let item: any = {};
                    let props = prop.complex.props;
                    for (let subProp of props) {
                        position++;
                        let value = getWorksheetValue(worksheet, `${numToColumn(position)}${i}`, subProp.actions);
                        if (subProp.checkNull && value === null) {
                            // NOTE: 
                            position += props.length - 1;
                            break;
                        }
                        item[subProp.name] = value;
                    }
                    // 自动注入index
                    item.index = j;
                    if (Object.getOwnPropertyNames(item).length > 0) {
                        info[prop.name].push(item);
                    }
                }
            }
            else {
                position++;
                // NOTE: 没看懂 先注释
                /* if (prop.offset) {
                    position += prop.offset;
                } */
                let value = getWorksheetValue(worksheet, `${numToColumn(position)}${i}`, prop.actions);
                if (prop.checkNull && value === null) {
                    // NOTE: 
                    position += props.length - 1;
                    break out;
                }
                info[prop.name] = value;
            }
        }
        configs.push(info);
    }
    return configs;
}

export function parseFile(filePath: string): void {
    let workbook = xlsx.readFile(filePath);
    let datas: Map<number, any[]> = new Map();
    // 遍历所有sheet,过滤掉名称不包含'-'
    for (let name of workbook.SheetNames) {
        let array = name.split("_");
        if (array.length > 1) {
            let type = parseInt(array[1].replace("type", ""));
            if (array[2]) {
                let stepId = parseInt(array[2]);
                const parseData = parse(workbook.Sheets[name], propMap[type])
                datas.set(stepId, parseData);
            }
        }
    }
    console.log({ datas })
    let configs: any[] = [];
    let worksheet = workbook.Sheets[workbook.SheetNames[0]];
    // 从第三行开始遍历
    for (let i = 3; ; i++) {
        let id = worksheet[`A${i}`];
        if (!id) {
            break;
        }
        // 获取环节类型
        let typeStr = getWorksheetValue(worksheet, `B${i}`);
        let type = parseInt(typeStr.split("_")[1].replace("type", ""));
        let stepId = getWorksheetValue(worksheet, `A${i}`);
        let step: any = {};
        // 环节stepId 默认就是number?
        step.stepId = stepId;
        step.stepType = type;
        if (type === 1) {
            step.cfg = {
                video: getWorksheetValue(worksheet, `C${i}`)
            };
        } else {
            let id = getWorksheetValue(worksheet, `C${i}`, [parseInt]);
            let quesDatas = datas.get(stepId);
            console.log(quesDatas)
            if (quesDatas) {
                for (let config of quesDatas) {
                    if (config.id === id) {
                        step.cfg = config;
                        break;
                    }
                }
            }
        }
        configs.push(step);
    }

    let dirName = path.dirname(filePath);
    let fileName = path.basename(filePath).replace(path.extname(filePath), "");
    // fileName = fileName.split("_")[1];
    fs.writeFileSync(path.join(dirName, `${fileName}.json`), JSON.stringify(configs));
}

function getWorksheetValue(worksheet: xlsx.WorkSheet, pos: string, actions: Function[] = [], defaultValue: any = null): any {
    if (worksheet[pos]) {
        let result = worksheet[pos].v;
        if (typeof result === 'string') {
            if (result.indexOf(".png.png") >= 0) {
                result = result.replace(".png.png", ".png");
            }
            if (result.indexOf("/bg.png") >= 0) {
                result = result.replace("bg.png", "bg.jpg");
            }
            result = result.trim();
        }
        if (actions && actions.length > 0) {
            for (let action of actions) {
                result = action(result);
            }
        }
        return result;
    }
    return defaultValue;
}

function splitToNumbers(raw: string): number[] {
    let result: number[] = [];
    if (raw && raw.length > 0) {
        for (let item of raw.split(",")) {
            result.push(Number(item));
        }
        return result;
    }
    return result;
}

function parseBoolean(raw: string): boolean {
    return raw === "是";
}

function numToColumn(raw: number) {
    const array: number[] = [];
    let numToString = function (nnum: number) {
        let num = nnum - 1;
        let a = Math.floor(num / 26);
        let b = num % 26;
        array.unshift(b + 'A'.charCodeAt(0)); // A charCode 65
        if (a > 0) {
            numToString(a);
        }
    }
    numToString(raw);
    // 转为字符串
    let char = '';
    for (let i = 0; i < array.length; i++) {
        char += String.fromCharCode(array[i]);
    }
    return char;
}

function subOne(raw: string | number): number | null {
    if (typeof raw === "number") {
        return raw - 1;
    }
    if (raw && raw.length > 0) {
        return parseInt(raw) - 1;
    }
    return null;
}

// parseFile("D:\\public\\excel配置表\\lesson4.xlsx");