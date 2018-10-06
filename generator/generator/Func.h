#pragma once
#include "stdafx.h"

using namespace std;


class matrix
{
private:
	bool o;
	int** matr;
	int v;
public:
	matrix(int ver, bool ori);

	void RandomMatrix() const;

	void PrintMatrix(ofstream fout) const;

	int GetVer() const;

	bool GetOri() const;

	~matrix();
};

int NameCheck(const char* name, const char* ext);

bool Generate(string conf, vector<matrix*> *matrixVector);
	

