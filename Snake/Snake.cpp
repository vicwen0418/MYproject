#include <iostream>
#include <stdlib.h>
#include <deque>                                                                    //雙向佇列
#include <windows.h>                                                                //平台短暫睡眠(讓遊戲不要跑那麼快)
#include <ctime>                                                                    //亂數
#include <conio.h>                                                                  //keyPress
using namespace std;
//目前已無gotoxy(),需自己實現
void gotoxy(short x, short y) {
    HANDLE hOut = GetStdHandle(STD_OUTPUT_HANDLE);                                  //Access控制程式碼
    COORD pos = { x, y };
    SetConsoleCursorPosition(hOut, pos);                                            //Move cursor to 指定位置
}
void HideCursor() { //Hide the cursor.
    CONSOLE_CURSOR_INFO cursor_info = { 1,0 };
    SetConsoleCursorInfo(GetStdHandle(STD_OUTPUT_HANDLE), &cursor_info);
}
struct Snakebody {                                                                  //snake's body.
    char body;
    short pos_x;
    short pos_y;                                                                    //Snake's position.
};
//遊戲所需素材
/*
                                                                                    map, score, food(eaten , xy), snake_body, speed, snake_state(alive or dead)
*/
class Game {
private:
    const int Map_width = 40;
    const int Map_height = 20;
    unsigned int score = 0;
    bool eaten_food = false;
    short food_pos_x, food_pos_y;
    deque<Snakebody> snake;
    unsigned int speed = 200;
    bool snake_state = true;
    char press = 'w';
public:
    //wall, putfood, judge_eatfood and state , keypress , generate_snake
    Game();
    void wall() {
        for (int i = 0; i < Map_height; i++) {
            gotoxy(0, i); //left
            cout << "■";
            gotoxy(Map_width, i);                                                   //right
            cout << "■";
        }
        for (int j = 0; j < Map_width; j += 2) {                                    // ***** 因■是兩位元 所以要+2
            gotoxy(j, 0);                                                           //top
            cout << "■";
            gotoxy(j, Map_height); //bottom
            cout << "■";
        }
    }
    void initial_snake() {                                                          //初始化蛇 直立型
        snake.push_back({ 'O', short(Map_width / 2), short(Map_height / 2) });      //新增蛇頭
        for (int i = 0; i < 2; i++)
            snake.push_back({ '#', short(Map_width / 2), short(Map_height / 2 + 1 + i) });
        snake.push_back({ ' ', short(Map_width / 2), short(Map_height / 2 + 4) }); //ADD tail，先放NULL，以便於後面新增節數時操作
    }
    void generate_snake() {                                                         //產生蛇的長相
        for (int k = 0; k < snake.size(); k++) {
            gotoxy(snake[k].pos_x, snake[k].pos_y);
            cout << snake[k].body;
        }
    }
    void put_food() {                                                               //放食物
        food_pos_x = rand() % (Map_width - 3) + 3;                                  //食物要在地圖中，不能在地圖邊框上。 ***
        food_pos_y = rand() % ((Map_height - 2) - 2) + 2;
        gotoxy(food_pos_x, food_pos_y);
        cout << 'x';
    }
    void judgement_eaten_food() {                                                   //判斷是否吃到食物
        if (food_pos_x == snake[0].pos_x && food_pos_y == snake[0].pos_y)
            eaten_food = true;
    }
    void judgement_snake_state() {                                                  //判斷是否撞牆了或撞到自己身體
        for (int i = 1; i < snake.size() - 1; i++) {
            if (snake[0].pos_x == snake[i].pos_x && snake[0].pos_y == snake[i].pos_y) {
                snake_state = false;
            }
        }
        if (snake[0].pos_x == 1 || snake[0].pos_y == 0 || snake[0].pos_x >= 40 || snake[0].pos_y >= 20)
            snake_state = false;
    }
    void KeyPress() {
        char key = press;
        if (_kbhit())                                                               //按下按鍵的函式
            press = _getch();
        for (int i = snake.size() - 1; i > 0; i--) {                                //蛇的移動方法，每移動一次，後面一節的身體到了它的前一節身體上
            snake[i].pos_x = snake[i - 1].pos_x;
            snake[i].pos_y = snake[i - 1].pos_y;
        }
        if (press == 'w' || press == 'W') {
            press = 'w';
            snake[0].pos_y--;
        }
        if (press == 's' || press == 'S') {
            press = 's';
            snake[0].pos_y++;
        }
        if (press == 'a' || press == 'A') {
            press = 'a';
            snake[0].pos_x--;
        }
        if (press == 'd' || press == 'D') {
            press = 'd';
            snake[0].pos_x++;
        }

    }
    void show() {                                                                    //show your score.
        gotoxy(50, 0);
        cout << "Your score : " << score << '\n';
    }
};
Game::Game() {
    HideCursor();
    srand(time(NULL));
    initial_snake();
    put_food();
    Snakebody tail;
    while (1) {
        wall();
        tail = snake.back();
        if (eaten_food == true) {
            snake.back().body = '#';
            snake.push_back(tail);
            gotoxy(food_pos_x, food_pos_y);
            cout << ' ';
            put_food();
            score++;
            eaten_food = false;
        }
        KeyPress();
        generate_snake();
        judgement_snake_state();
        if (snake_state == false)
            break;
        judgement_eaten_food();
        Sleep(speed);
        show();

    }
}
int main()
{
    system("color 70");
    system("title 貪食蛇(Snake)");
    Game Game;
    gotoxy(15, 22);
    cout << "GAME OVER." << '\n';
}
