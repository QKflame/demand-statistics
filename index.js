import xlsx from 'node-xlsx';
import chalk from 'chalk';
import {table} from 'table';
import _ from 'lodash';
import path from 'path';
import {fileURLToPath} from 'url';
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const sheets = xlsx.parse(`${__dirname}/data.xlsx`);

console.log('\n');

const greenProgress = ['已上线', '已完成'];
const redProgress = ['hold', '待完善', 'pending', '待开发', '待开始'];
const blueProgress = ['待测试', '测试中', '已提测', '待上线', '等待qa排期测试'];
const yellowProgress = ['开发中', '待联调', '联调中'];

function calcColumnProportion(num1, num2) {
    if (!num1) {
        return '-';
    }
    return `${num1} (${((num1 * 100) / num2).toFixed(2)}%)`;
}

function calcProportion(num1, num2) {
    if (!num1) {
        return '-';
    }
    return `${((num1 * 100) / num2).toFixed(2)}%`;
}

function renderProgress(progress) {
    return greenProgress.includes(progress)
        ? chalk.green(progress)
        : redProgress.includes(progress)
        ? chalk.red(progress)
        : blueProgress.includes(progress)
        ? chalk.blue(progress)
        : yellowProgress.includes(progress)
        ? chalk.cyan(progress)
        : progress;
}

function genDaysProgressTableData(daysProgressMapInPlan, daysProgressMapOutPlan, daysInPlan, daysOutPlan) {
    let keys = _.uniq(_.concat(Object.keys(daysProgressMapInPlan), Object.keys(daysProgressMapOutPlan)));
    let data = [[chalk.bold('进展'), chalk.bold('规划内 (占比)'), chalk.bold('规划外 (占比)')]];
    const getCompareNum = val => {
        let ret = -1;
        [greenProgress, blueProgress, yellowProgress, redProgress].some((item, index) => {
            if (item.includes(val)) {
                ret = index;
                return true;
            }
        });
        return ret;
    };
    keys = keys.sort((a, b) => {
        const aNum = getCompareNum(a);
        const bNum = getCompareNum(b);
        return aNum - bNum;
    });
    keys.forEach(key => {
        data.push([
            renderProgress(key),
            calcColumnProportion(daysProgressMapInPlan[key], daysInPlan),
            calcColumnProportion(daysProgressMapOutPlan[key], daysOutPlan)
        ]);
    });
    return data;
}

function percentageToDecimal(percentage) {
    const percentageValue = parseFloat(percentage.replace('%', ''));
    const decimalValue = percentageValue / 100;
    return decimalValue;
}

let totalDaysInPlan = 0;
let totalDaysOutPlan = 0;
let totalDaysInSupport = 0;
let productSummaryMap = [];

const priorities = [0, 1, 2, -1];

function calculateObjectSum(obj) {
    let sum = 0;

    for (const key in obj) {
        if (typeof obj[key] === 'number') {
            // 如果属性值是数字，将其添加到总和中
            sum += obj[key];
        } else if (typeof obj[key] === 'object') {
            // 如果属性值是对象，递归调用函数以计算子对象的总和
            sum += calculateObjectSum(obj[key]);
        }
    }

    return sum;
}

function setSheetNamePriorityProgressDaysMap(map, sheetName, inOrOut, priority, progress, days) {
    const path = `${sheetName}.${inOrOut}.${priority}.${progress}.days`;
    let val = _.get(map, path);
    if (!val) {
        val = 0;
    }

    _.set(map, path, val + days);
}

function getPriorityDaysSummaryTableData(map = {}) {
    const getDays = (inOrOut, priority) => {
        return calculateObjectSum(_.get(map, `${inOrOut}.${priority}`) || {});
    };

    const getOnlinePercent = (inOrOut, priority) => {
        const days = getDays(inOrOut, priority);
        let onlineDays = 0;
        Object.keys(_.get(map, `${inOrOut}.${priority}`) || {}).forEach(progress => {
            if (greenProgress.includes(progress)) {
                onlineDays += _.get(map, `${inOrOut}.${priority}.${progress}.days`) || 0;
            }
        });
        let ret = onlineDays / days;
        let percent = ret == 0 || isNaN(ret) ? '-' : (ret * 100).toFixed(2) + '%';
        return ret === '-'
            ? ret
            : ret >= 0.8 || inOrOut === 'out'
            ? chalk.green.bold(percent)
            : chalk.red.bold(percent);
    };

    const getDevPercent = (inOrOut, priority) => {
        const days = getDays(inOrOut, priority);
        let devDays = 0;
        Object.keys(_.get(map, `${inOrOut}.${priority}`) || {}).forEach(progress => {
            if (greenProgress.includes(progress) || blueProgress.includes(progress)) {
                devDays += _.get(map, `${inOrOut}.${priority}.${progress}.days`) || 0;
            }
        });
        let ret = devDays / days;
        let percent = ret == 0 || isNaN(ret) ? '-' : (ret * 100).toFixed(2) + '%';
        return ret === '-'
            ? ret
            : ret >= 0.8 || inOrOut === 'out'
            ? chalk.green.bold(percent)
            : chalk.red.bold(percent);
    };

    return priorities.map(item => {
        const priority = `P${item}`;
        return [
            priority === 'P-1' ? '无' : priority,
            getDays('in', priority),
            getOnlinePercent('in', priority),
            getDevPercent('in', priority),
            getDays('out', priority),
            getOnlinePercent('out', priority),
            getDevPercent('out', priority)
        ];
    });
}

// 根据优先级，统计人力
let sheetNamePriorityProgressDaysMap = {};

sheets.forEach(({name: sheetName, data: sheetData}) => {
    let daysInPlan = 0;
    let daysOutPlan = 0;
    let finishedOnlineDays = 0;
    let finishedDevDays = 0;
    let daysProgressMapInPlan = {};
    let daysProgressMapOutPlan = {};
    let redItems = [];
    let blueItems = [];
    let greenItems = [];

    sheetData
        .filter(([title]) => title)
        .forEach(([title, priority, personInCharge, days, progress]) => {
            days = parseFloat(days);

            progress = (progress || '待完善')
                .toLowerCase()
                .replace(/(\d+%)/g, '')
                .trim();

            priority = (priority || 'P-1').trim().toUpperCase().replace(/\s/g, '');

            if (!_.concat(greenProgress, blueProgress, yellowProgress, redProgress).includes(progress)) {
                throw new Error(`The "${progress}" progress value is invalid!`);
            }

            if (!isNaN(days)) {
                const isOutPlan = /规划外/.test(title);

                setSheetNamePriorityProgressDaysMap(
                    sheetNamePriorityProgressDaysMap,
                    sheetName,
                    isOutPlan ? 'out' : 'in',
                    priority,
                    progress,
                    days
                );

                if (isOutPlan) {
                    daysOutPlan += days;

                    if (!daysProgressMapOutPlan[progress]) {
                        daysProgressMapOutPlan[progress] = 0;
                    }

                    daysProgressMapOutPlan[progress] += days;
                } else {
                    daysInPlan += days;

                    if (!daysProgressMapInPlan[progress]) {
                        daysProgressMapInPlan[progress] = 0;
                    }

                    daysProgressMapInPlan[progress] += days;
                }
            }

            if (greenProgress.includes(progress)) {
                finishedOnlineDays += days;
            }

            if (greenProgress.includes(progress) || blueProgress.includes(progress)) {
                finishedDevDays += days;
            }

            if (redProgress.includes(progress) && title && title.trim()) {
                redItems.push({
                    title,
                    priority,
                    personInCharge,
                    days,
                    progress
                });
            }

            if ((blueProgress.includes(progress) || yellowProgress.includes(progress)) && title && title.trim()) {
                blueItems.push({
                    title,
                    priority,
                    personInCharge,
                    days,
                    progress
                });
            }

            if (greenProgress.includes(progress) && title && title.trim()) {
                greenItems.push({
                    title,
                    priority,
                    personInCharge,
                    days,
                    progress
                });
            }

            if (!redProgress.includes(progress)) {
                totalDaysInSupport += days;
            }
        });

    console.log(chalk.white.bgBlue.bold(`${sheetName}\n`));

    totalDaysInPlan += daysInPlan;
    totalDaysOutPlan += daysOutPlan;

    const totalDays = daysInPlan + daysOutPlan;
    console.log(`${chalk.bold('需求人力汇总:')}`);
    console.log(
        table([
            [chalk.bold('统计'), chalk.bold('规划内'), chalk.bold('规划外'), chalk.bold('总人力')],
            ['汇总', daysInPlan, daysOutPlan, totalDays],
            [
                '占比',
                `${((daysInPlan * 100) / totalDays).toFixed(2)}%`,
                percentageToDecimal(`${((daysOutPlan * 100) / totalDays).toFixed(2)}%`) >= 0.2
                    ? chalk.red.bold(`${((daysOutPlan * 100) / totalDays).toFixed(2)}%`)
                    : `${((daysOutPlan * 100) / totalDays).toFixed(2)}%`,
                '-'
            ]
        ])
    );

    console.log(`${chalk.bold('需求进展汇总:')}`);
    console.log(
        table(genDaysProgressTableData(daysProgressMapInPlan, daysProgressMapOutPlan, daysInPlan, daysOutPlan))
    );

    if (greenItems.length) {
        console.log(`${chalk.bold('已完成需求:\n')}`);
        console.log(
            `${greenItems
                .map(item => {
                    return `${item.title} ${item.priority ? `【${item.priority.toUpperCase()}】` : ''} ${chalk.bold(
                        item.personInCharge
                    )} ${item.days ? chalk.red.bold(item.days) + '人天' : ''} ${renderProgress(item.progress)}`;
                })
                .join('\n')}\n`
        );
    }

    if (redItems.length) {
        console.log(`${chalk.bold('未开始需求:\n')}`);
        console.log(
            `${redItems
                .map(item => {
                    return `${item.title} ${item.priority ? `【${item.priority.toUpperCase()}】` : ''} ${chalk.bold(
                        item.personInCharge
                    )} ${item.days ? chalk.red.bold(item.days) + '人天' : ''} ${renderProgress(item.progress)}`;
                })
                .join('\n')}\n`
        );
    }

    if (blueItems.length) {
        console.log(`${chalk.bold('进行中需求:\n')}`);
        console.log(
            `${blueItems
                .map(item => {
                    return `${item.title} ${item.priority ? `【${item.priority.toUpperCase()}】` : ''} ${chalk.bold(
                        item.personInCharge
                    )} ${item.days ? chalk.red.bold(item.days) + '人天' : ''} ${renderProgress(item.progress)}`;
                })
                .join('\n')}\n`
        );
    }

    // 需求优先级分析
    console.log(`${chalk.bold('需求完成情况:')}`);
    console.log(
        table([
            [
                '优先级',
                '人力(规划内)',
                '上线率(规划内)',
                '研发完成率(规划内)',
                '人力(规划外)',
                '上线率(规划外)',
                '研发完成率(规划外)'
            ].map(item => chalk.bold(item)),
            ...getPriorityDaysSummaryTableData(sheetNamePriorityProgressDaysMap[sheetName])
        ])
    );

    // 上线率: 已上线的需求（规划内 + 规划外） / 规划内人力
    let finishedOnlinePercent = calcProportion(finishedOnlineDays, daysInPlan);
    console.log(
        `${chalk.bold('总体需求上线率:')} ${
            percentageToDecimal(finishedOnlinePercent) >= 0.8
                ? chalk.green(finishedOnlinePercent)
                : chalk.red(finishedOnlinePercent)
        }\n`
    );

    // 研发完成率： 已研发完成的需求（规划内 + 规划外）/ 规划内人力
    let finishedDevPercent = calcProportion(finishedDevDays, daysInPlan);
    console.log(
        `${chalk.bold('总体需求研发完成率:')} ${
            percentageToDecimal(finishedDevPercent) >= 0.8
                ? chalk.green(finishedDevPercent)
                : chalk.red(finishedDevPercent)
        }\n`
    );

    productSummaryMap.push({
        summaryName: sheetName,
        finishedOnlinePercent,
        finishedDevPercent
    });

    console.log('============================================================');
    console.log('\n');
});

console.log(chalk.white.bgBlue.bold(`各产品合并统计\n`));
console.log(`${chalk.bold('人力汇总:')}`);
const totalDays = totalDaysInPlan + totalDaysOutPlan;
console.log(
    table([
        [chalk.bold('统计'), chalk.bold('规划内'), chalk.bold('规划外'), chalk.bold('总人力')],
        ['汇总', totalDaysInPlan, totalDaysOutPlan, totalDays],
        [
            '占比',
            `${((totalDaysInPlan * 100) / totalDays).toFixed(2)}%`,
            percentageToDecimal(`${((totalDaysOutPlan * 100) / totalDays).toFixed(2)}%`) >= 0.2
                ? chalk.red.bold(`${((totalDaysOutPlan * 100) / totalDays).toFixed(2)}%`)
                : `${((totalDaysOutPlan * 100) / totalDays).toFixed(2)}%`,
            '-'
        ]
    ])
);

console.log(`${chalk.bold('需求统计:')}`);
console.log(
    table([
        [chalk.bold('产品'), chalk.bold('上线率'), chalk.bold('研发完成率')],
        ..._(
            productSummaryMap.map(item => {
                return [
                    item.summaryName,
                    percentageToDecimal(item.finishedOnlinePercent) >= 0.8
                        ? chalk.green.bold(item.finishedOnlinePercent)
                        : chalk.red.bold(item.finishedOnlinePercent),
                    percentageToDecimal(item.finishedDevPercent) >= 0.8
                        ? chalk.green.bold(item.finishedDevPercent)
                        : chalk.red.bold(item.finishedDevPercent)
                ];
            })
        )
    ])
);

const totalDaysInQuarter = 320;
console.log(chalk.bold('本季度 FE 人力约: ', totalDaysInQuarter + ' 人天\n'));

console.log(
    `${chalk.bold(
        '本季度已支持的需求人力约:',
        totalDaysInSupport >= totalDaysInQuarter
            ? chalk.green.bold(totalDaysInSupport)
            : chalk.red.bold(totalDaysInSupport),
        '人天'
    )}`
);
