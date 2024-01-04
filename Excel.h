#pragma once
#include <string>
#include <iostream>
#include <stdexcept>
#include <vector>
#include <fstream>
#include <windows.h>
using namespace std;

class Excel
{
public:
	class Cell
	{
	public:
		Cell* left;
		Cell* right;
		Cell* up;
		Cell* down;
		string data;
		string color;
		string align;

		Cell()
		{
			left = nullptr;
			right = nullptr;
			up = nullptr;
			down = nullptr;
			data = "";
		}
	};

	Cell* current;
	Cell* root;

	Excel()
	{
		Cell* nRoot = new Cell();
		current = nRoot;
		root = nRoot;
		for (int i = 1; i < 5; i++)
		{
			Cell* newCell = new Cell();
			newCell->left = current;
			current->right = newCell;
			current = current->right;
		}

		current = root;
		Cell* temp = current;
		for (int i = 0; i < 4; i++)
		{
			insertRowBelow();
			temp = temp->down;
			current = temp;
		}
		current = root;
	}

	void insertRowAbove()
	{
		Cell* prev = nullptr;
		while (current->left != nullptr)
		{
			current = current->left;
		}

		if (current == root)
		{
			while (current != nullptr)
			{
				Cell* newCell = new Cell();
				newCell->down = current;
				if (current->up != nullptr)
				{
					newCell->up = current->up;
					current->up->down = newCell;
				}
				current->up = newCell;

				if (prev != nullptr)
				{
					prev->right = newCell;
					newCell->left = prev;
				}
				prev = newCell;
				current = current->right;
			}

			current = root;
			root = root->up;
		}

		else
		{
			while (current != nullptr)
			{
				Cell* newCell = new Cell();
				newCell->down = current;
				if (current->up != nullptr)
				{
					newCell->up = current->up;
					current->up->down = newCell;
				}
				current->up = newCell;

				if (prev != nullptr)
				{
					prev->right = newCell;
					newCell->left = prev;
				}
				prev = newCell;
				current = current->right;
			}
		}
	}

	void insertRowBelow()
	{
		Cell* prev = nullptr;
		while (current->left != nullptr)
		{
			current = current->left;
		}

		while (current != nullptr)
		{
			Cell* newCell = new Cell();
			newCell->up = current;
			if (current->down != nullptr)
			{
				newCell->down = current->down;
				current->down->up = newCell;
			}
			current->down = newCell;

			if (prev != nullptr)
			{
				prev->right = newCell;
				newCell->left = prev;
			}
			prev = newCell;
			current = current->right;
		}
		current = root;
	}

	void insertColumnToRight()
	{
		Cell* prev = nullptr;

		while (current->up != nullptr)
		{
			current = current->up;
		}

		if (current->right == nullptr)
		{
			while (current != nullptr)
			{
				Cell* newCell = new Cell();
				newCell->left = current;
				current->right = newCell;

				if (prev != nullptr)
				{
					newCell->up = prev;
					prev->down = newCell;
				}

				prev = newCell;
				current = current->down;
			}
		}

		else
		{
			while (current != nullptr)
			{
				Cell* newCell = new Cell();
				newCell->left = current;
				newCell->right = current->right;
				current->right->left = newCell;
				current->right = newCell;
				if (prev != nullptr)
				{
					prev->down = newCell;
					newCell->up = prev;
				}
				prev = newCell;
				current = current->down;
			}
		}

		current = root;
	}

	void insertColumnToLeft()
	{
		Cell* prev = nullptr;

		while (current->up != nullptr)
		{
			current = current->up;
		}

		if (current->left == nullptr)
		{
			while (current != nullptr)
			{
				Cell* newCell = new Cell();
				newCell->right = current;
				current->left = newCell;
				if (prev != nullptr)
				{
					prev->down = newCell;
					newCell->up = prev;
				}

				prev = newCell;
				current = current->down;
			}
			root = root->left;
		}

		else
		{
			while (current != nullptr)
			{
				Cell* newCell = new Cell();
				newCell->right = current;
				newCell->left = current->left;
				current->left->right = newCell;
				current->left = newCell;

				if (prev != nullptr)
				{
					prev->down = newCell;
					newCell->up = prev;
				}
				prev = newCell;
				current = current->down;
			}
		}

		current = root;
	}

	void insertCellByRightShift()
	{
		Cell* curr = current;
		string prevData;

		while (current->right != nullptr)
		{
			current = current->right;
		}

		insertColumnToRight();
		current = curr;
		prevData = curr->data;
		curr->data = "";
		current = current->right;
		while (current != nullptr)
		{
			string temp = current->data;
			current->data = prevData;
			prevData = temp;
			current = current->right;
		}
		current = root;
	}

	void insertCellByDownShift()
	{
		Cell* curr = current;
		string prevData;

		while (current->down != nullptr)
		{
			current = current->down;
		}

		insertRowBelow();
		current = curr;
		prevData = curr->data;
		curr->data = "";
		current = current->down;
		while (current != nullptr)
		{
			string temp = current->data;
			current->data = prevData;
			prevData = temp;
			current = current->down;
		}
		current = root;
	}

	void deleteCellByLeftShift()
	{
		Cell* curr = current;
		while (current->right != nullptr)
		{
			current->data = current->right->data;
			current = current->right;
		}
		current->data = "";
		current = root;
	}

	void deleteCellByUpShift()
	{
		Cell* curr = current;
		while (current->down != nullptr)
		{
			current->data = current->down->data;
			current = current->down;
		}
		current->data = "";
		current = root;
	}

	void deleteColumn()
	{
		if (current == nullptr)
		{
			return;
		}

		else
		{
			while (current->up != nullptr)
			{
				current = current->up;
			}

			if (current->right != nullptr && current->left != nullptr)
			{
				while (current != nullptr)
				{
					Cell* temp = current;
					current->left->right = current->right;
					current->right->left = current->left;
					current = current->down;
					delete temp;
				}
			}

			else if (current->right == nullptr && current->left != nullptr)
			{
				while (current != nullptr)
				{
					Cell* temp = current;
					current->left->right = nullptr;
					current = current->down;
					delete temp;
				}
			}

			else
			{
				root = root->right;
				while (current != nullptr)
				{
					Cell* temp = current;
					current->right->left = nullptr;
					current = current->down;
					delete temp;
				}

			}
			current = root;
		}
	}

	void deleteRow()
	{
		if (current == nullptr)
		{
			return;
		}

		else
		{
			while (current->left != nullptr)
			{
				current = current->left;
			}

			if (current->up != nullptr && current->down != nullptr)
			{
				while (current != nullptr)
				{
					Cell* temp = current;
					current->up->down = current->down;
					current->down->up = current->up;
					current = current->right;
					delete temp;
				}
			}

			else if (current->down == nullptr && current->up != nullptr)
			{
				while (current != nullptr)
				{
					Cell* temp = current;
					current->up->down = nullptr;
					current = current->right;
					delete temp;
				}
			}

			else
			{
				root = root->down;
				while (current != nullptr)
				{
					Cell* temp = current;
					current->down->up = nullptr;
					current = current->right;
					delete temp;
				}
			}
			current = root;
		}
	}

	void clearRow()
	{
		Cell* temp = current;
		while (temp->left != nullptr)
		{
			temp = temp->left;
		}

		while (temp != nullptr)
		{
			temp->data = "";
			temp = temp->right;
		}
	}

	void clearColumn()
	{
		Cell* temp = current;
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}

		while (temp != nullptr)
		{
			temp->data = "";
			temp = temp->down;
		}
		current = root;
	}

	class Iterator
	{
		Cell* current;
	public:
		Iterator(Cell* cell)
		{
			current = cell;
		}

		Iterator& operator ++()
		{
			current = current->down;
		}

		Iterator& operator --()
		{
			current = current->up;
		}

		Iterator& operator ++(int)
		{
			current = current->right;
		}

		Iterator& operator --(int)
		{
			current = current->left;
		}

		std::string& operator *()
		{
			return current->data;
		}
	};

	float getRangeSumResult(Cell* rangeStart, Cell* rangeEnd)
	{
		float sum = 0;
		bool isRangeCol = false;
		Cell* temp = rangeStart;
		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return 0;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				isRangeCol = true;
				break;
			}
			temp = temp->down;
		}

		temp = rangeStart;
		if (isRangeCol)
		{
			while (temp != rangeEnd->down)
			{
				try
				{
					sum += stof(temp->data);
					temp = temp->down;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}

		else
		{
			while (temp != rangeEnd->right)
			{
				try
				{
					sum += stof(temp->data);
					temp = temp->right;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}
		return sum;
	}

	float getRangeAverageResult(Cell* rangeStart, Cell* rangeEnd)
	{
		float sum = 0;
		int totalValues = 0;
		bool isRangeCol = false;
		Cell* temp = rangeStart;
		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return 0;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				isRangeCol = true;
				break;
			}
			temp = temp->down;
		}

		temp = rangeStart;
		if (isRangeCol)
		{
			while (temp != rangeEnd->down)
			{
				try
				{
					sum += stof(temp->data);
					temp = temp->down;
					totalValues++;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}

		else
		{
			while (temp != rangeEnd->right)
			{
				try
				{
					sum += stof(temp->data);
					temp = temp->right;
					totalValues++;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}
		return sum / totalValues;
	}

	float getRangeSum(Cell* rangeStart, Cell* rangeEnd)
	{
		float sum = 0;
		bool isRangeCol = false;
		Cell* updatingCell = current;
		Cell* temp = rangeStart;
		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return 0;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				isRangeCol = true;
				break;
			}
			temp = temp->down;
		}

		temp = rangeStart;
		if (isRangeCol)
		{
			while (temp != rangeEnd->down)
			{
				try
				{
					sum += stof(temp->data);
					temp = temp->down;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}

		else
		{
			while (temp != rangeEnd->right)
			{
				try
				{
					sum += stof(temp->data);
					temp = temp->right;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}
		updatingCell->data = "" + to_string(sum);
		return sum;
	}

	float getRangeAverage(Cell* rangeStart, Cell* rangeEnd)
	{
		float sum = 0;
		int totalValues = 0;
		bool isRangeCol = false;
		Cell* updatingCell = current;
		Cell* temp = rangeStart;
		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return 0;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				isRangeCol = true;
				break;
			}
			temp = temp->down;
		}

		temp = rangeStart;
		if (isRangeCol)
		{
			while (temp != rangeEnd->down)
			{
				try
				{
					sum += stof(temp->data);
					temp = temp->down;
					totalValues++;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}

		else
		{
			while (temp != rangeEnd->right)
			{
				try
				{
					sum += stof(temp->data);
					temp = temp->right;
					totalValues++;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}
		updatingCell->data = "" + to_string(sum / totalValues);
		return sum / totalValues;
	}

	int getRangeCount(Cell* rangeStart, Cell* rangeEnd)
	{
		int totalValues = 0;
		bool isRangeCol = false;
		Cell* temp = rangeStart;
		Cell* updatingCell = current;

		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return 0;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				isRangeCol = true;
				break;
			}
			temp = temp->down;
		}

		temp = rangeStart;
		if (isRangeCol)
		{
			while (temp != rangeEnd->down)
			{
				try
				{
					temp = temp->down;
					totalValues++;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}

		else
		{
			while (temp != rangeEnd->right)
			{
				try
				{
					temp = temp->right;
					totalValues++;
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}
		updatingCell->data = "" + to_string(totalValues);

		return totalValues;
	}

	float getRangeMinimum(Cell* rangeStart, Cell* rangeEnd)
	{
		float minimum = INT_MAX;
		bool isRangeCol = false;
		Cell* temp = rangeStart;
		Cell* updatingCell = current;

		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return 0;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				isRangeCol = true;
				break;
			}
			temp = temp->down;
		}

		temp = rangeStart;
		if (isRangeCol)
		{
			while (temp != rangeEnd->down)
			{
				try
				{
					float curr_data = stof(temp->data);
					temp = temp->down;

					if (curr_data < minimum)
					{
						minimum = curr_data;
					}
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}

		else
		{
			while (temp != rangeEnd->right)
			{
				try
				{
					float curr_data = stof(temp->data);
					temp = temp->right;

					if (curr_data < minimum)
					{
						minimum = curr_data;
					}
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}
		updatingCell->data = "" + to_string(minimum);
		return minimum;
	}

	float getRangeMaximum(Cell* rangeStart, Cell* rangeEnd)
	{
		float maximum = INT_MIN;
		bool isRangeCol = false;
		Cell* updatingCell = current;
		Cell* temp = rangeStart;

		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return 0;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				isRangeCol = true;
				break;
			}
			temp = temp->down;
		}

		temp = rangeStart;
		if (isRangeCol)
		{
			while (temp != rangeEnd->down)
			{
				try
				{
					float curr_data = stof(temp->data);
					temp = temp->down;

					if (curr_data > maximum)
					{
						maximum = curr_data;
					}
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}

		else
		{
			while (temp != rangeEnd->right)
			{
				try
				{
					float curr_data = stof(temp->data);
					temp = temp->right;

					if (curr_data > maximum)
					{
						maximum = curr_data;
					}
				}

				catch (const invalid_argument& e)
				{
					return 0;
				}
			}
		}
		updatingCell->data = "" + to_string(maximum);
		return maximum;
	}

	vector<string> copy(Cell* rangeStart, Cell* rangeEnd)
	{
		bool isRangeCol = false;
		Cell* temp = rangeStart;
		vector<string> result;
		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return result;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				isRangeCol = true;
				break;
			}
			temp = temp->down;
		}

		temp = rangeStart;
		if (isRangeCol)
		{
			while (temp != rangeEnd->down)
			{
				try
				{
					result.push_back(temp->data);
					temp = temp->down;
				}

				catch (const invalid_argument& e)
				{
					return result;
				}
			}
		}

		else
		{
			while (temp != rangeEnd->right)
			{
				try
				{
					result.push_back(temp->data);
					temp = temp->right;
				}

				catch (const invalid_argument& e)
				{
					return result;
				}
			}
		}
		return result;
	}

	vector<string> cut(Cell* rangeStart, Cell* rangeEnd)
	{
		bool isRangeCol = false;
		Cell* temp = rangeStart;
		vector<string> result;
		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return result;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				isRangeCol = true;
				break;
			}
			temp = temp->down;
		}

		temp = rangeStart;
		if (isRangeCol)
		{
			while (temp != rangeEnd->down)
			{
				try
				{
					result.push_back(temp->data);
					temp->data = "";
					temp = temp->down;
				}

				catch (const invalid_argument& e)
				{
					return result;
				}
			}
		}

		else
		{
			while (temp != rangeEnd->right)
			{
				try
				{
					result.push_back(temp->data);
					temp->data = "";
					temp = temp->right;
				}

				catch (const invalid_argument& e)
				{
					return result;
				}
			}

		}
		return result;
	}

	bool isColumn(Cell* rangeStart, Cell* rangeEnd)
	{
		Cell* temp = rangeStart;
		if (rangeStart == nullptr || rangeEnd == nullptr)
		{
			return false;
		}

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				return true;
			}
			temp = temp->down;
		}

		temp = rangeStart;

		while (temp != nullptr)
		{
			if (temp == rangeEnd)
			{
				return false;
			}
			temp = temp->right;
		}
	}

	void paste(vector<string> data, bool isColumn)
	{
		Cell* updating = current;
		int size = data.size();
		if (isColumn)
		{
			while (current->up != nullptr)
			{
				current = current->up;
			}

			for (int i = 0; i < size; i++)
			{
				updating = current;
				current->data = data[i];
				if (current->down == nullptr && i < size - 1)
				{
					insertRowBelow();
					current = updating;
				}
				current = current->down;
			}
		}

		else
		{
			while (current->left != nullptr)
			{
				current = current->left;
			}

			for (int i = 0; i < size; i++)
			{
				updating = current;
				current->data = data[i];
				if (current->right == nullptr && i < size - 1)
				{
					insertColumnToRight();
					current = updating;
				}
				current = current->right;
			}
		}
		current = root;
	}

	void saveSheet(string filename)
	{
		Cell* start = root;
		Cell* temp = root;
		current = root;
		fstream file;
		file.open(filename, ios::out);


		while (temp != nullptr)
		{
			while (temp != nullptr)
			{
				current = temp;
				file << current->data;

				if (temp->right != nullptr)
				{
					file << ",";
				}
				temp = temp->right;
			}
			file << ",eol";
			if (start->down != nullptr)
			{
				file << "\n";
			}
			start = start->down;
			temp = start;
		}

		file.close();
	}

	void loadSheet(string filename)
	{
		Cell* start = root;
		string line;
		fstream file;
		Cell* curr = start;
		file.open(filename, ios::in);
		current = root;
		while (getline(file, line))
		{
			int field = 0;
			string currWord;
			start = current;
			curr = current;
			if (current->down == nullptr)
			{
				insertRowBelow();
				current = curr;
			}

			while (currWord != "eol")
			{
				currWord = parseData(line, field);
				if (currWord == "eol")
				{
					break;
				}
				else if (current->right == nullptr && parseData(line, field + 1) != "eol")
				{
					curr = current;
					insertColumnToRight();
					current = curr;
				}
				current->data = currWord;
				field++;
				current = current->right;

			}
			current = start;
			current = current->down;
		}
		file.close();
		current = root;
	}

	string parseData(string line, int field)
	{
		int commaCount = 0;
		string reqdWord;
		for (int idx = 0; idx < line.length(); idx++)
		{
			if (line[idx] == ',')
			{
				commaCount++;
			}
			else if (commaCount == field)
			{
				reqdWord = reqdWord + line[idx];
			}
		}
		return reqdWord;
	}

	void display()
	{
		Cell* temp = root;
		while (temp != nullptr)
		{
			Cell* prev = temp;
			while (temp != nullptr)
			{
				if (temp->data != "")
				{
					cout << temp->data << " --> ";
				}
				else
				{
					cout << "! --> ";
					if (temp->right == nullptr)
					{
						cout << "nextNull";
					}
				}
				temp = temp->right;
			}
			cout << endl;
			if (prev->down == nullptr)
			{
				cout << "nextNull" << endl;
			}
			temp = prev->down;
		}
	}
};

class FrontEnd
{
	Excel::Cell* rootRow;
	Excel::Cell* rootCol;
	public:
	int x = 2;
	int y = 2;
	FrontEnd(Excel::Cell* node)
	{
		rootRow = node;
		rootCol = node;
		x = 2;
		y = 2;
	}

	void displayGrid(Excel::Cell* node)
	{
		x = 2;
		y = 2;
		Excel::Cell* tempR = node;
		Excel::Cell* tempC = node;
		while (tempR != nullptr)
		{
			tempC = tempR;
			while (tempC != nullptr)
			{
				printCell(tempC, x, y);
				tempC = tempC->right;
				x += 10;
			}
			y += 4;
			x = 2;

			tempR = tempR->down;
		}
		x = 2;
		y = 2;
	}

	void printCell(Excel::Cell* cell, int x, int y)
	{
		gotoxy(x, y);
		cout << "\033[36m" << "----------";
		gotoxy(x, y + 1);
		cout << "|        |";
		gotoxy(x, y + 2);
		cout << "|        |";
		gotoxy(x + 3, y + 2);
		cout << cell->data;
		gotoxy(x, y + 3);
		cout << "|        |";
		gotoxy(x, y + 4);
		cout << "----------";
	}

	void showCurrentCell(Excel::Cell* current, int x, int y)
	{
		gotoxy(x, y);
		cout << "\033[32m" << "----------";
		gotoxy(x, y + 1);
		cout << "|        |";
		gotoxy(x, y + 2);
		cout << "|        |";
		gotoxy(x + 3, y + 2);
		cout << current->data;
		gotoxy(x, y + 3);
		cout << "|        |";
		gotoxy(x, y + 4);
		cout << "----------";
		gotoxy(x + 3, y + 2);
	}

	void printDataInCell(Excel::Cell* cell)
	{
		gotoxy(x + 3, y + 2);
		cout << cell->data;
	}
	
	string input()
	{
		string userInput;
		string userIn;
		getline(cin, userIn);
		for (int i = 0; i < 4; i++)
		{
			if (userIn[i] == '\0' || userIn[i] == '\n')
			{
				userInput[i] = userIn[i];
				break;
			}
			else
			{
				userInput += userIn[i];
			}
		}
		return userInput;
	}

	void gotoxy(int x, int y)
	{
		COORD coordinates;
		coordinates.X = x;
		coordinates.Y = y;
		SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
	}

	void printSumAvg(Excel::Cell* root, double sum, double avg)
	{
		y = 2;
		Excel::Cell* temp = root;
		while (temp != nullptr)
		{
			y += 4;
			temp = temp->down;
		}

		gotoxy(5, y + 4);
		cout << "Sum: " << sum << "      Average: " << avg;
		y = 2;
	}
};