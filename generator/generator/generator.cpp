// generator.cpp: определяет точку входа для консольного приложения.
//

#include "stdafx.h"
#include "Func.h"

using namespace std;


int main(int argc, char** argv)
{
	if (argc != 5)
	{
		cout << "Wrong number of parameters" << endl;
		return -1;
	}

	char* config;
	char* output;

	if (!strcmp(argv[1], "--config"))
	{
		config = argv[2];
	}
	else
	{
		cout << "Wrong first key, --config expected" << endl;
		return -2;
	}

	if (!strcmp(argv[3], "--output"))
	{
		output = argv[4];
		int i = 0;
		while (output[i])
		{
			if (output[i] == '.')
			{
				cout << "Wrong second key, --output expected" << endl;
				return -2;
			}
			i++;
		}
	}
	else
	{
		cout << "Wrong second key, --output expected" << endl;
		return -2;
	}

	char ext[] = ".cfg";
	if (NameCheck(config, ext))
	{
		cout << "Wrong extension of config, .cfg expected" << endl;
		return -4;
	}

	fstream fin;

	fin.open(config, ios::in);
	if (!fin.is_open())
	{
		cout << "Can't open config file" << endl;
		return -4;
	}
	vector<string> numVec;
	vector<matrix*> rezVec;
	while (!fin.eof())
	{
		string hstr;
		getline(fin, hstr);
		numVec.push_back(hstr);
	}
	const int file_num = numVec.size();
	fin.close();
	for (int i = 0; i < file_num; i++)
	{
		string hstr;
		Generate(numVec[i], &rezVec);
	}

	

	for(matrix* s : rezVec)
	{
		//s->PrintMatrix();
		const int len = 255;
		char buff[len];
		string oFileName = "";
		oFileName += output;
		_itoa_s(s->GetVer(), buff, 10);
		oFileName += buff;
		if(s->GetOri())
		{
			oFileName += "o";
		}
		else
		{
			oFileName += "n";
		}
		oFileName += ".cfg";
		ofstream fout;
		fout.open(oFileName, ios::out);
		s->PrintMatrix(move(fout));
		fout.close();
		cout << oFileName << endl;
		delete s;
	}
	return 0;
}
