#include <cstdlib>
#include <iostream>

using namespace std;

int main(int argc, char *argv[])
{
    short num;
	FILE *fp = fopen("../������.dat", "rb+"); 
	
	if( fp == NULL )
	{
        printf("������ ������ �����ϴ�.\n");
		printf("������ ������ ��ġ�� ���� ������ ���� �������� �մϴ�.\n");
        return EXIT_SUCCESS;
	} 
	
	fread( &num, 2, 1, fp );
	printf("���� ���������� %d�Դϴ�.\n", num );
	printf("���ο� ���������� ���ϼ���. ( 0 = ��� ) : " );
	scanf("%d", &num);
	
	if( num == 0 ) 
	    printf("��� �Ǿ����ϴ�.\n"); 
	else
	{
	 	fseek( fp, 0, SEEK_SET );
	    fwrite( &num, 2, 1, fp );
	    printf("���� ������ ��ȯ�߽��ϴ�.\n"); 
    }
    fclose(fp);
  	return EXIT_SUCCESS;
}
