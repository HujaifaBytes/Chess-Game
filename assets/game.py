import math
import random
import os
import urllib.request
import pygame
from pygame import mixer

def download_file(url, filename):
    urllib.request.urlretrieve(url, filename)

# Create temp directory for assets
if not os.path.exists('assets'):
    os.makedirs('assets')
pygame.init()

screen = pygame.display.set_mode((800, 600))


try:
    background = pygame.image.load('assets/background.png')
except pygame.error:
    background = pygame.Surface((800, 600))
    background.fill((0, 0, 0))


try:
    mixer.music.load('assets/bgmusic.wav')
    mixer.music.play(-1)
except pygame.error:
    print("Warning: Background music could not be loaded")


pygame.display.set_caption("Space Invader")
icon = pygame.image.load('assets/ufo.png')
pygame.display.set_icon(icon)

# Player
playerImg = pygame.image.load('assets/player.png')
playerX = 370
playerY = 480
playerX_change = 0


enemyImg = []
enemyX = []
enemyY = []
enemyX_change = []
enemyY_change = []
num_of_enemies = 6

for i in range(num_of_enemies):
    enemyImg.append(pygame.image.load('assets/enemy.png'))
    enemyX.append(random.randint(0, 736))
    enemyY.append(random.randint(50, 150))
    enemyX_change.append(4)
    enemyY_change.append(40)

# Bullet
bulletImg = pygame.image.load('assets/bullet.png')
bulletX = 0
bulletY = 480
bulletX_change = 0
bulletY_change = 10
bullet_state = "ready"

# Score
score_value = 0
font = pygame.font.Font(None, 32)

textX = 10
testY = 10

# Game Over
over_font = pygame.font.Font(None, 64)

def show_score(x, y):
    score = font.render("Score : " + str(score_value), True, (255, 255, 255))
    screen.blit(score, (x, y))

def game_over_text():
    over_text = over_font.render("GAME OVER", True, (255, 255, 255))
    screen.blit(over_text, (200, 250))

def player(x, y):
    screen.blit(playerImg, (x, y))

def enemy(x, y, i):
    screen.blit(enemyImg[i], (x, y))

def fire_bullet(x, y):
    global bullet_state
    bullet_state = "fire"
    screen.blit(bulletImg, (x + 16, y + 10))

def isCollision(enemyX, enemyY, bulletX, bulletY):
    distance = math.sqrt(math.pow(enemyX - bulletX, 2) + (math.pow(enemyY - bulletY, 2)))
    return distance < 27

# Game Loop
running = True
game_over = False

while running:
    screen.fill((0, 0, 0))
    screen.blit(background, (0, 0))
    
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False

        if not game_over:
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_LEFT:
                    playerX_change = -5
                if event.key == pygame.K_RIGHT:
                    playerX_change = 5
                if event.key == pygame.K_SPACE:
                    if bullet_state == "ready":
                        try:
                            bulletSound = mixer.Sound('assets/laser.wav')
                            bulletSound.play()
                        except pygame.error:
                            print("Warning: Laser sound could not be played")
                        bulletX = playerX
                        fire_bullet(bulletX, bulletY)

            if event.type == pygame.KEYUP:
                if event.key == pygame.K_LEFT or event.key == pygame.K_RIGHT:
                    playerX_change = 0
        else:
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:
                    running = False

    if not game_over:
        playerX += playerX_change
        playerX = max(0, min(playerX, 736))

        # Enemy Movement
        for i in range(num_of_enemies):
            # Game Over
            if enemyY[i] > 440:
                game_over = True
                break

            enemyX[i] += enemyX_change[i]
            if enemyX[i] <= 0:
                enemyX_change[i] = 4
                enemyY[i] += enemyY_change[i]
            elif enemyX[i] >= 736:
                enemyX_change[i] = -4
                enemyY[i] += enemyY_change[i]

            # Collision
            if isCollision(enemyX[i], enemyY[i], bulletX, bulletY):
                try:
                    explosionSound = mixer.Sound('assets/explosion.wav')
                    explosionSound.play()
                except pygame.error:
                    print("Warning: Explosion sound could not be played")
                bulletY = 480
                bullet_state = "ready"
                score_value += 1
                enemyX[i] = random.randint(0, 736)
                enemyY[i] = random.randint(50, 150)

            enemy(enemyX[i], enemyY[i], i)

        # Bullet Movement
        if bulletY <= 0:
            bulletY = 480
            bullet_state = "ready"

        if bullet_state == "fire":
            fire_bullet(bulletX, bulletY)
            bulletY -= bulletY_change

        player(playerX, playerY)
        show_score(textX, testY)
    else:
        game_over_text()

    pygame.display.update()
    pygame.time.delay(5)


pygame.quit()