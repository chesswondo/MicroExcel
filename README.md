# MicroExcel
Лабораторна робота №1, варіант 7, ООП, курс 2, семестр 1

Має бути головне меню та кнопки, що дублюють основні його елементи. Повинна
бути подана коротка інформація про програму та надаватися допомога користувачу.
Діалог україномовний.

Під "аналізом та обчисленням" тут розуміється клас задач, що часто зустрічаються
в системному програмуванні у процесі обробки інформації операційними системами,
макроасемблерами, асемблерами, трансляторами. Усі вони передбачають завдання
вхідної інформації (виразів) із використанням формалізованих описів (БНФ, діаграм
Вірта та ін.). Перший етап визначає синтаксичну правильність виразів, подальші -
перетворення формату даних, заключний - обчислення значення виразу.

Варіант роботи видається викладачем.

Розширити функціональність файлового менеджера шляхом додавання форм для
введення, обробки та збереження електронних таблиць.

Клітини електронної таблиці містять вирази, що складаються із знаків операцій,
констант та посилань на інші клітинки таблиці. Синтаксис посилання на клітину
таблиці може бути запропонований виконавцем роботи за аналогією з відомими
системами. Перевірити синтаксичну правильність виразу та знайти його значення
(передбачити виконання двох етапів: синтаксична перевірка з локалізацією помилок та
власне обчислення значення). Відображення інформації в таблиці на формі має
підтримувати два режими: ВИРАЗ/ЗНАЧЕННЯ.

Для додаткових двох балів замість виразів у клітинках таблиці - оператори
спрощеної мови програмування.

Вважати, що при побудові виразів використовуються:

1. цілі числа довільної довжини
2. круглі дужки
3. імена клітинок (напр., А3)
4. операції та функції, які для кожного із варіантів лабораторної роботи
визначаються окремо.

Символи пропуску у виразах використовуються звичайним чином. Імена клітинок
відділяються від чисел та інших імен пробілом.

Обов'язковим є включення до звіту формалізованого опису множини виразів за
допомогою БНФ або діаграм Вірта.

Перелік варіантів операцій та функцій:

    +, -, *, / (бінарні операції);  
    
    mod, dіv;  
    
    ^ (піднесення у степінь);  
    
    mmax(x1,x2,...,xN),  
    mmіn(x1,x2,...,xN) (N>=1);

вираз вважати арифметичним (значення такого виразу повинно бути числом)

