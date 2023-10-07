import xlsx from 'node-xlsx';
import chalk from 'chalk';
import {table} from 'table';
import _ from 'lodash';
import path from 'path';
import {fileURLToPath} from 'url';
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const sheets = xlsx.parse(`${__dirname}/day.xlsx`);

const dataMap = {};

const P0 = 'P0';
const P1 = 'P1';
const P2 = 'P2';

const priorities = [P0, P1, P2];

sheets.forEach(({name, data}) => {
    let p0Summary = 0;
    let p1Summary = 0;
    let p2Summary = 0;
    data.forEach(([priority, days]) => {
        if (_.isEmpty(priority && priority.trim()) && _.isEmpty(days && days.trim())) {
            return;
        }

        priority = priority && priority.toUpperCase().trim();

        if (!priorities.includes(priority)) {
            throw new Error('The priority info is invalid!');
        }

        days = parseFloat(days);
        if (isNaN(days)) {
            throw new Error('The days info is invalid!');
        }

        if (priority === P0) {
            p0Summary += days;
        } else if (priority === P1) {
            p1Summary += days;
        } else if (priority === P2) {
            p2Summary += days;
        }
    });

    _.set(dataMap, `${name}.p0Summary`, p0Summary);
    _.set(dataMap, `${name}.p1Summary`, p1Summary);
    _.set(dataMap, `${name}.p2Summary`, p2Summary);
});

let tableData = [];

let p0TotalSummary = 0;
let p1TotalSummary = 0;
let p2TotalSummary = 0;
let totalSummary = 0;

Object.keys(dataMap).forEach(key => {
    const p0Summary = _.get(dataMap, `${key}.p0Summary`);
    const p1Summary = _.get(dataMap, `${key}.p1Summary`);
    const p2Summary = _.get(dataMap, `${key}.p2Summary`);

    p0TotalSummary += p0Summary;
    p1TotalSummary += p1Summary;
    p2TotalSummary += p2Summary;

    const total = p0Summary + p1Summary + p2Summary;
    totalSummary += total;
    tableData.push([key, p0Summary, p1Summary, p2Summary, total]);
});

console.log(
    table([
        ['产品', 'P0需求', 'P1需求', 'P2需求', '汇总'].map(item => chalk.bold(item)),
        ...tableData,
        [chalk.bold('汇总'), p0TotalSummary, p1TotalSummary, p2TotalSummary, totalSummary]
    ])
);

const perDays = 62;
const totalDays = perDays * (1 + 1 + 1 + 0.5 + 0.5 * 3);

console.log(`${chalk.bold('本季度 FE 人力约:')} ${chalk.bold(totalDays)}人日\n`);

const diff = totalSummary - totalDays;
if (diff > 0) {
    console.log(chalk.bold(`预计本季度人力缺口为: ${chalk.red((diff / perDays).toFixed(2))} 人力`));
} else {
    console.log(chalk.bold(`预计本季度人力富余为: ${chalk.green((-diff / perDays).toFixed(2))} 人力`));
}
