#include <iostream>
#include <bitset>
#include <locale>
using namespace std;

int main() {
    setlocale(LC_ALL, "ru");   

    float x;
    cout << "Введите число: ";
    cin >> x;

    uint32_t bits = *reinterpret_cast<uint32_t*>(&x);
    bitset<32> b(bits);

    cout << "IEEE 754 (32-bit): " << b << endl;
    cout << "Знак:        " << b[31] << endl;
    cout << "Порядок:     " << b.to_string().substr(1, 8) << endl;
    cout << "Мантисса:    " << b.to_string().substr(9) << endl;

    return 0;
}
