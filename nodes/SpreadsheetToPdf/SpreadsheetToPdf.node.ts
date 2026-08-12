import { spawn } from 'child_process';
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from 'fs/promises';
import { tmpdir } from 'os';
import { basename, extname, join, parse } from 'path';
import { pathToFileURL } from 'url';

import type {
	IBinaryData,
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { ApplicationError, NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';

const SUPPORTED_EXTENSIONS = new Set(['.csv', '.ods', '.xls', '.xlsm', '.xlsx']);
const DEFAULT_TIMEOUT_SECONDS = 60;
const MAX_PROCESS_OUTPUT_CHARS = 12000;

type ProcessResult = {
	exitCode: number | null;
	signal: NodeJS.Signals | null;
	stdout: string;
	stderr: string;
};

type SourceFile = {
	fileName: string;
	extension: string;
};

export class SpreadsheetToPdf implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Spreadsheet to PDF',
		name: 'spreadsheetToPdf',
		icon: 'file:spreadsheetToPdf.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{"Convert " + $parameter["inputBinaryPropertyName"] + " to PDF"}}',
		description: 'Convert spreadsheet binary data to PDF using LibreOffice',
		defaults: {
			name: 'Spreadsheet to PDF',
		},
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		usableAsTool: true,
		properties: [
			{
				displayName: 'Input Binary Property',
				name: 'inputBinaryPropertyName',
				type: 'string',
				default: 'data',
				required: true,
				description: 'Name of the binary property containing the spreadsheet file',
			},
			{
				displayName: 'Output Binary Property',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				required: true,
				description: 'Name of the binary property where the generated PDF will be written',
			},
			{
				displayName: 'Output File Name',
				name: 'outputFileName',
				type: 'string',
				default: '',
				placeholder: 'order.pdf',
				description: 'Optional PDF filename. If omitted, the input filename is reused with a .pdf extension.',
			},
			{
				displayName: 'Timeout (Seconds)',
				name: 'timeoutSeconds',
				type: 'number',
				default: DEFAULT_TIMEOUT_SECONDS,
				typeOptions: {
					minValue: 1,
				},
				description: 'Maximum time to allow LibreOffice to run for each item',
			},
			{
				displayName: 'LibreOffice Executable',
				name: 'libreOfficeExecutable',
				type: 'string',
				default: '',
				placeholder: 'soffice',
				description:
					'Optional path or command for LibreOffice. Leave empty to auto-detect soffice or libreoffice from PATH.',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				const inputBinaryPropertyName = this.getNodeParameter(
					'inputBinaryPropertyName',
					itemIndex,
				) as string;
				const outputBinaryPropertyName = this.getNodeParameter(
					'outputBinaryPropertyName',
					itemIndex,
				) as string;
				const outputFileNameParameter = this.getNodeParameter('outputFileName', itemIndex, '') as string;
				const timeoutSeconds = this.getNodeParameter(
					'timeoutSeconds',
					itemIndex,
					DEFAULT_TIMEOUT_SECONDS,
				) as number;
				const libreOfficeExecutableParameter = this.getNodeParameter(
					'libreOfficeExecutable',
					itemIndex,
					'',
				) as string;

				const item = items[itemIndex];
				const binaryData = item.binary?.[inputBinaryPropertyName];

				if (!binaryData) {
					throw new NodeOperationError(
						this.getNode(),
						`No binary data found on property "${inputBinaryPropertyName}"`,
						{ itemIndex },
					);
				}

				const sourceFile = getSourceFile(binaryData, inputBinaryPropertyName);

				if (!SUPPORTED_EXTENSIONS.has(sourceFile.extension)) {
					throw new NodeOperationError(
						this.getNode(),
						`Unsupported spreadsheet extension "${sourceFile.extension || '(none)'}". Supported extensions: ${[
							...SUPPORTED_EXTENSIONS,
						].join(', ')}`,
						{ itemIndex },
					);
				}

				const timeoutMs = Math.max(1, timeoutSeconds) * 1000;
				const executable = await resolveLibreOfficeExecutable(
					libreOfficeExecutableParameter.trim(),
					Math.min(timeoutMs, 5000),
				);
				const inputBuffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
				const conversionStart = Date.now();
				const conversion = await convertSpreadsheetToPdf(
					executable,
					sourceFile.fileName,
					inputBuffer,
					timeoutMs,
				);
				const elapsedMs = Date.now() - conversionStart;
				const outputFileName = normalizePdfFileName(
					outputFileNameParameter,
					`${parse(sourceFile.fileName).name}.pdf`,
				);
				const outputBinaryData = await this.helpers.prepareBinaryData(
					conversion.pdfBuffer,
					outputFileName,
					'application/pdf',
				);
				const warnings = getWarnings(conversion.result);
				const metadata: IDataObject = {
					sourceFileName: sourceFile.fileName,
					sourceMimeType: binaryData.mimeType ?? null,
					sourceBinaryProperty: inputBinaryPropertyName,
					outputFileName,
					outputBinaryProperty: outputBinaryPropertyName,
					mimeType: 'application/pdf',
					executable,
					command: {
						executable,
						args: conversion.args,
					},
					elapsedMs,
				};

				if (warnings.length > 0) {
					metadata.warnings = warnings;
				}

				returnData.push({
					json: {
						...item.json,
						xlsxToPdf: metadata,
					},
					binary: {
						...item.binary,
						[outputBinaryPropertyName]: outputBinaryData,
					},
					pairedItem: { item: itemIndex },
				});
			} catch (error) {
				if (this.continueOnFail()) {
					returnData.push({
						json: {
							...items[itemIndex].json,
							error: getErrorMessage(error),
						},
						pairedItem: { item: itemIndex },
					});
					continue;
				}

				if (error instanceof NodeOperationError) {
					throw new NodeOperationError(this.getNode(), error, { itemIndex });
				}

				throw new NodeOperationError(this.getNode(), getErrorMessage(error), { itemIndex });
			}
		}

		return [returnData];
	}
}

async function resolveLibreOfficeExecutable(
	configuredExecutable: string,
	timeoutMs: number,
): Promise<string> {
	if (configuredExecutable) {
		await verifyExecutable(configuredExecutable, timeoutMs);
		return configuredExecutable;
	}

	for (const candidate of [
		'soffice',
		'libreoffice',
		'C:\\Program Files\\LibreOffice\\program\\soffice.com',
		'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
		'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com',
		'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
	]) {
		try {
			await verifyExecutable(candidate, timeoutMs);
			return candidate;
		} catch (error) {
			if (!isExecutableMissingError(error)) {
				continue;
			}
		}
	}

	throw new ApplicationError(
		'LibreOffice executable was not found. Install LibreOffice in the n8n runtime image/container, or set the LibreOffice Executable parameter to the soffice/libreoffice path.',
	);
}

async function verifyExecutable(executable: string, timeoutMs: number): Promise<void> {
	const result = await runProcess(executable, ['--version'], timeoutMs);

	if (result.exitCode !== 0) {
		throw new ApplicationError(
			`LibreOffice executable check failed for "${executable}" with exit code ${result.exitCode ?? 'unknown'}. ${summarizeProcessOutput(
				result,
			)}`,
		);
	}
}

async function convertSpreadsheetToPdf(
	executable: string,
	sourceFileName: string,
	inputBuffer: Buffer,
	timeoutMs: number,
): Promise<{ args: string[]; pdfBuffer: Buffer; result: ProcessResult }> {
	const workDir = await mkdtemp(join(tmpdir(), 'n8n-xlsx-to-pdf-'));
	const profileDir = join(workDir, 'lo-profile');
	const inputPath = join(workDir, sourceFileName);

	try {
		await mkdir(profileDir);
		await writeFile(inputPath, inputBuffer);

		const args = [
			`-env:UserInstallation=${pathToFileURL(profileDir).href}`,
			'--headless',
			'--nologo',
			'--nofirststartwizard',
			'--nodefault',
			'--nolockcheck',
			'--norestore',
			'--convert-to',
			'pdf',
			'--outdir',
			workDir,
			inputPath,
		];
		const result = await runProcess(executable, args, timeoutMs);

		if (result.exitCode !== 0) {
			throw new ApplicationError(
				`LibreOffice failed with exit code ${result.exitCode ?? 'unknown'}. ${summarizeProcessOutput(
					result,
				)}`,
			);
		}

		const pdfPath = await findOutputPdfPath(workDir, sourceFileName);

		if (!pdfPath) {
			throw new ApplicationError(
				`LibreOffice did not create a PDF output file. ${summarizeProcessOutput(result)}`,
			);
		}

		return {
			args,
			pdfBuffer: await readFile(pdfPath),
			result,
		};
	} finally {
		await rm(workDir, { recursive: true, force: true });
	}
}

async function findOutputPdfPath(workDir: string, sourceFileName: string): Promise<string | undefined> {
	const expectedPdfPath = join(workDir, `${parse(sourceFileName).name}.pdf`);

	try {
		await readFile(expectedPdfPath);
		return expectedPdfPath;
	} catch {
		const files = await readdir(workDir);
		const pdfFile = files.find((file) => extname(file).toLowerCase() === '.pdf');

		return pdfFile ? join(workDir, pdfFile) : undefined;
	}
}

function getSourceFile(binaryData: IBinaryData, inputBinaryPropertyName: string): SourceFile {
	const inferredExtension = getExtensionFromMimeType(binaryData.mimeType);
	const fallbackFileName = `${inputBinaryPropertyName}${inferredExtension || '.xlsx'}`;
	const fileName = sanitizeFileName(binaryData.fileName || fallbackFileName);
	const extension = extname(fileName).toLowerCase();

	return {
		fileName,
		extension,
	};
}

function getExtensionFromMimeType(mimeType: string | undefined): string {
	switch (mimeType) {
		case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
			return '.xlsx';
		case 'application/vnd.ms-excel':
			return '.xls';
		case 'application/vnd.ms-excel.sheet.macroenabled.12':
		case 'application/vnd.ms-excel.sheet.macroEnabled.12':
			return '.xlsm';
		case 'text/csv':
			return '.csv';
		case 'application/vnd.oasis.opendocument.spreadsheet':
			return '.ods';
		default:
			return '';
	}
}

function normalizePdfFileName(fileName: string, fallbackFileName: string): string {
	const normalizedFileName = sanitizeFileName(fileName || fallbackFileName);

	if (extname(normalizedFileName).toLowerCase() === '.pdf') {
		return normalizedFileName;
	}

	return `${normalizedFileName}.pdf`;
}

function sanitizeFileName(fileName: string): string {
	const sanitized = [...basename(fileName)]
		.map((character) => {
			if (character.charCodeAt(0) < 32 || '<>:"/\\|?*'.includes(character)) {
				return '_';
			}

			return character;
		})
		.join('')
		.trim();

	return sanitized || 'spreadsheet.xlsx';
}

function runProcess(executable: string, args: string[], timeoutMs: number): Promise<ProcessResult> {
	return new Promise((resolve, reject) => {
		// This node must launch LibreOffice to perform the spreadsheet conversion.
		// eslint-disable-next-line @n8n/community-nodes/no-dangerous-functions -- LibreOffice is the required conversion engine.
		const childProcess = spawn(executable, args, { windowsHide: true });
		let stdout = '';
		let stderr = '';
		let timedOut = false;
		let settled = false;
		const timeout = setTimeout(() => {
			timedOut = true;
			childProcess.kill('SIGKILL');
		}, timeoutMs);

		childProcess.stdout?.on('data', (chunk: Buffer) => {
			stdout = appendProcessOutput(stdout, chunk);
		});

		childProcess.stderr?.on('data', (chunk: Buffer) => {
			stderr = appendProcessOutput(stderr, chunk);
		});

		childProcess.on('error', (error) => {
			clearTimeout(timeout);
			settled = true;
			reject(error);
		});

		childProcess.on('close', (exitCode, signal) => {
			clearTimeout(timeout);

			if (settled) {
				return;
			}

			settled = true;

			if (timedOut) {
				reject(
					new ApplicationError(
						`LibreOffice timed out after ${Math.round(timeoutMs / 1000)} seconds`,
					),
				);
				return;
			}

			resolve({
				exitCode,
				signal,
				stdout,
				stderr,
			});
		});
	});
}

function appendProcessOutput(currentOutput: string, chunk: Buffer): string {
	const nextOutput = currentOutput + chunk.toString('utf8');

	if (nextOutput.length <= MAX_PROCESS_OUTPUT_CHARS) {
		return nextOutput;
	}

	return nextOutput.slice(nextOutput.length - MAX_PROCESS_OUTPUT_CHARS);
}

function isExecutableMissingError(error: unknown): boolean {
	return typeof error === 'object' && error !== null && 'code' in error && error.code === 'ENOENT';
}

function summarizeProcessOutput(result: ProcessResult): string {
	const stdout = result.stdout.trim();
	const stderr = result.stderr.trim();
	const parts = [];

	if (stdout) {
		parts.push(`stdout: ${stdout}`);
	}

	if (stderr) {
		parts.push(`stderr: ${stderr}`);
	}

	return parts.join(' ') || 'No output was written by LibreOffice.';
}

function getWarnings(result: ProcessResult): string[] {
	const stderr = result.stderr.trim();

	return stderr ? [stderr] : [];
}

function getErrorMessage(error: unknown): string {
	if (error instanceof Error) {
		return error.message;
	}

	return String(error);
}
