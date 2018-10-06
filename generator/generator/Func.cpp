#include "stdafx.h"
#include "Func.h"


using namespace std;

int NameCheck(const char* name, const char* ext)
{
	int i = 0;
	int n = 0;
	while (name[i] != '\0')
	{
		i++;
	}
	while (ext[n] != '\0')
	{
		n++;
	}
	if (i < n + 1)
	{
		return -1;
	}
	for (int k = n; k > 0; k--)
	{
		if (name[i - k] != ext[n-k])
		{
			return -1;
		}
	}

	return 0;
}

matrix::matrix(int ver, bool ori)
{
	this->o = ori;
	this->v = ver;
	matr = new int*[ver];
	for (int i = 0; i < ver; i++)
	{
		matr[i] = new int[ver];
	}
	for (int i = 0; i < ver; i++)
	{
		for (int k = 0; k < ver; k++)
		{
			matr[i][k] = 0;
		}
	}
}
void matrix::RandomMatrix() const
{
	if (this->o == 0)
	{
		for (int i = 0; i < this->v; i++)
		{
			for (int k = i + 1; k < this->v; k++)
			{
				srand(time(0) + k);
				this->matr[i][k] = (rand() % 1000 + 1) % 2;
				this->matr[k][i] = matr[i][k];
			}
		}
		for (int i = 0; i < this->v; i++)
		{
			int counter = 0;
			for (int k = 0; k < this->v; k++)
			{
				if (this->matr[i][k] == 1)
				{
					counter++;
				}
			}
			if (counter < 1)
			{
				srand(time(0) + i);
				int k = i;
				while (k == i)
				{
					k = rand() % v + 0;
				}
				matr[i][k] = 1;
				matr[k][i] = 1;
			}
		}
		return;
	}
	for (int i = 0; i < this->v; i++)
	{
		for (int k = i + 1; k < this->v; k++)
		{
			srand(time(0) + k);
			this->matr[i][k] = ((rand() % 3 + 1) - 2);
			this->matr[k][i] = 0 - matr[i][k];
		}
	}
	for (int i = 0; i < this->v; i++)
	{
		int counter = 0;
		for (int k = 0; k < this->v; k++)
		{
			if (this->matr[i][k] != 0)
			{
				counter++;
			}
		}
		if (counter < 1)
		{
			srand(time(0) + i);
			int k = i;
			while (k == i)
			{
				k = rand() % v + 0;
			}
			const int rez = rand() % 1 + 0;
			if (rez)
			{
				matr[i][k] = -1;
				matr[k][i] = 1;
			}
			matr[i][k] = 1;
			matr[k][i] = -1;
		}
	}
}

void matrix::PrintMatrix(ofstream fout) const
{
	for (int i = 0; i < this->v; i++)
	{
		for (int k = 0; k < this->v; k++)
		{
			fout << matr[i][k] << " ";
		}
		fout << endl;
	}
}

int matrix::GetVer() const
{
	return v;
}

bool matrix::GetOri() const
{
	return o;
}

matrix::~matrix()
{
	for (int i = 0; i < v; i++)
	{
		delete[]matr[i];
	}
}

bool Generate(string conf, vector<matrix*>* matrixVector)
{
	string orient = "";
	string num = "";
	int i = 0;
	while (conf[i] != ' ')
	{
		if (!isdigit(conf[i]))
		{
			return false;
		}
		orient += conf[i];
		i++;
	}
	i++;
	while (conf[i])
	{
		if (!isdigit(conf[i]))
		{
			return false;
		}
		num += conf[i];
		i++;
	}
	if (orient == "0")
	{
		const bool ori = false;
		const int ver = stoi(num);
		auto mat = new matrix(ver, ori);
		mat->RandomMatrix();
		matrixVector->push_back(mat);
		return true;
	}
	if (orient == "1")
	{
		const bool ori = true;
		const int ver = stoi(num);
		auto mat = new matrix(ver, ori);
		mat->RandomMatrix();
		matrixVector->push_back(mat);
		return true;
	}
	return false;
}

