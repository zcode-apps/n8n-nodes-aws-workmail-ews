module.exports = {
	root: true,
	env: {
		node: true,
	},
	parser: '@typescript-eslint/parser',
	parserOptions: {
		ecmaVersion: 2020,
		sourceType: 'module',
		project: './tsconfig.json',
	},
	plugins: ['@typescript-eslint'],
	extends: [
		'eslint:recommended',
		'plugin:@typescript-eslint/recommended',
	],
	rules: {
		'@typescript-eslint/no-explicit-any': 'warn', // Warnung statt Fehler für bestehenden Code
		'@typescript-eslint/no-unused-vars': ['error', {
			argsIgnorePattern: '^_',
			varsIgnorePattern: '^_',
		}],
		'no-console': 'warn', // Warnung statt Fehler für Debugging
		'no-eval': 'error',
		'no-implied-eval': 'error',
	},
};
