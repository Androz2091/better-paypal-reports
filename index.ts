import xlsx from 'node-xlsx';
import chalk from 'chalk';
import { argv } from 'yargs';

const fileName = argv.file as string;

if (!fileName) {
    console.error('Please specify a valid file to process!');
    process.exit(1);
}

const daysInMonth = (month: number, year: number) => new Date(year, month, 0).getDate();
const USDToEuro = (n: number) => n * 0.75;
const space = (count: number) => console.log('\n'.repeat(count-1));

const id = fileName.split('-')[1];
const date = fileName.split('-')[2];
const year = date.substr(0, 4);
const month = date.substr(4, 2);
const day = date.substr(6, 2);

const parsed = xlsx.parse(fileName);
const rows = parsed[0].data;
rows.splice(0, 2);
rows.splice(rows.length-1, rows.length);

const data = rows.map((e) => {
    return {
        ID: e[0],
        item: e[1],
        client: e[2],
        status: e[3],
        price: USDToEuro(parseInt((e[4] as string).split(' ')[0], 10)),
        sold: e[5],
        date: e[6] as Date
    };
});

const activeSubscriptionCount = data.length;
const gainTotal = data.map((o) => o.price).reduce((p, c) => p + c);

const invoicesForThisMonth = data.filter((o) => `${new Date(o.date).getMonth()}.${new Date(o.date).getFullYear()}` === `${parseInt(month)-1}.${year}`);
const gainForThisMonth = invoicesForThisMonth.map((o) => o.price).reduce((p, c) => p + c);
const remainingDays = daysInMonth(parseInt(month), parseInt(year)) - parseInt(day);

const invoicesCreateByclientsWithMultipleInvoices = data.filter((o) => data.filter((o2) => o2.client === o.client).length > 1);
const clientsWithMultipleInvoices = Array.from(new Set(data.filter((o) => data.filter((o2) => o2.client === o.client).length > 1).map((o) => o.client)));
const gainSameClients = invoicesCreateByclientsWithMultipleInvoices.map((o) => o.price).reduce((p, c) => p + c);

space(1);
console.log(`${chalk.green("PAYPAL REPORT ")} ${id} (${chalk.yellow(`${day}/${month}/${year}`)})`);
space(2);
console.log(`Abonnements actifs: ${chalk.bold(chalk.green(activeSubscriptionCount))} (${chalk.bold(chalk.blue(`${gainTotal}€`))} par mois)`);
console.log(`Paiements attendus avant la fin du mois (${chalk.bold(remainingDays)} jours): ${chalk.bold(chalk.green(invoicesForThisMonth.length))} (${chalk.bold(chalk.cyan(`${(invoicesForThisMonth.length * 100 / activeSubscriptionCount).toFixed()}%`))} des abonnements). (${chalk.bold(chalk.blue(`${gainForThisMonth}€`))})`);
space(1);
console.log(`Clients qui payent pour plusieurs abonnements: ${chalk.bold(chalk.magenta(clientsWithMultipleInvoices.length))}. Cela représente ${chalk.bold(chalk.green(`${invoicesCreateByclientsWithMultipleInvoices.length}`))} abonnements (${chalk.bold(chalk.cyan(`${(invoicesCreateByclientsWithMultipleInvoices.length * 100 / activeSubscriptionCount).toFixed()}%`))} des abonnements). (${chalk.bold(chalk.blue(`${gainSameClients}€`))})`);
