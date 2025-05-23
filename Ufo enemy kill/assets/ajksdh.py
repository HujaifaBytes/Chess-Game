def game_loop():
    global player_x_change, bullet_state, bullet_x, bullet_y, score_value, lives
    
    reset_game_state()
    mixer.music.play(-1)
    
    start_time = pygame.time.get_ticks()

    running = True
    while running:
        screen.blit(background_img, (0, 0))
        
        # --- Increase difficulty over time ---
        if (pygame.time.get_ticks() - start_time) > 60000: # Every 60 seconds
            globals()['enemy_speed_multiplier'] += 0.2
            start_time = pygame.time.get_ticks()

        # --- Event Handling ---
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                save_to_excel("Quit")
                pygame.quit()
                exit()
            
            if event.type == pygame.MOUSEBUTTONDOWN:
                if quit_button_rect.collidepoint(event.pos):
                    save_to_excel("Quit")
                    running = False # Exit game loop, return to home

            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_LEFT: player_x_change = -sensitivity
                if event.key == pygame.K_RIGHT: player_x_change = sensitivity
                if event.key == pygame.K_SPACE and bullet_state == "ready":
                    bullet_sound.play()
                    bullet_state = "fire"
                    bullet_x, bullet_y = player_x, player_y
            
            if event.type == pygame.KEYUP:
                if event.key in [pygame.K_LEFT, pygame.K_RIGHT]:
                    player_x_change = 0
                    
        # --- Player Movement ---
        player_x += player_x_change
        player_x = max(0, min(player_x, SCREEN_WIDTH - 64))

        # Enemy movement and collision
        for enemy in list(enemies):
            enemy["x"] += enemy["x_change"]
            if enemy["x"] <= 0 or enemy["x"] >= SCREEN_WIDTH - 64:
                enemy["x_change"] *= -1
                enemy["y"] += 40
                
        if bullet_state == "fire" and math.sqrt((enemy["x"] - bullet_x)**2 + (enemy["y"] - bullet_y)**2) < 27:
                explosion_sound.play()
                bullet_state = "ready"
                score_value += 1
                enemies.remove(enemy)
                enemies.append(create_enemy())
            
        # --- Game Over ---
        if enemy["y"] > 480:
            lives -= 1
            enemies.remove(enemy)
            enemies.append(create_enemy())
            if lives <= 0:
                running = False

        if bullet_state == "fire":
            bullet_y -= bullet_y_change
            if bullet_y <= 0: bullet_state = "ready"

        # --- Drawing / Rendering ---
        if bullet_state == "fire":
            screen.blit(bullet_img, (bullet_x + 16, bullet_y + 10))
        
        screen.blit(player_img, (player_x, player_y))
        
        for enemy in enemies:
            screen.blit(enemy_img, (enemy['x'], enemy['y']))
        
        show_game_ui()
        pygame.display.update()
        
        
        
        
        
        #Home screen`
        def home_screen():
    global user_name, user_gender
    name_box = pygame.Rect(SCREEN_WIDTH/2 - 150, 200, 300, 50)
    male_box = pygame.Rect(SCREEN_WIDTH/2 - 150, 300, 140, 50)
    female_box = pygame.Rect(SCREEN_WIDTH/2 + 10, 300, 140, 50)
    start_box = pygame.Rect(SCREEN_WIDTH/2 - 100, 400, 200, 50)
    
    active_color = pygame.Color('lightskyblue3')
    passive_color = pygame.Color('gray15')
    name_color = passive_color
    
    name_active = False
    user_name = ""
    user_gender = ""

    while True:
        screen.blit(background_img, (0, 0))
        draw_text("Space Invaders Deluxe", title_font, WHITE, screen, SCREEN_WIDTH / 2, 80)
        draw_text("Enter Your Name:", info_font, WHITE, screen, SCREEN_WIDTH / 2, 170)
        
        mouse_pos = pygame.mouse.get_pos()
        
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                exit()
            if event.type == pygame.MOUSEBUTTONDOWN:
                if name_box.collidepoint(event.pos):
                    name_active = True
                else:
                    name_active = False
                
                if male_box.collidepoint(event.pos):
                    user_gender = "Male"
                elif female_box.collidepoint(event.pos):
                    user_gender = "Female"

                if start_box.collidepoint(event.pos) and user_name and user_gender:
                    return # Exit home screen and start the game
            
            if event.type == pygame.KEYDOWN and name_active:
                if event.key == pygame.K_BACKSPACE:
                    user_name = user_name[:-1]
                else:
                    user_name += event.unicode

        name_color = active_color if name_active else passive_color
        
        # Draw input boxes
        pygame.draw.rect(screen, name_color, name_box, 2)
        draw_button(male_box, 'Male', BLUE if user_gender == "Male" else passive_color)
        draw_button(female_box, 'Female', (255,105,180) if user_gender == "Female" else passive_color)
        
        # Draw Start button (only enable if name and gender are set)
        if user_name and user_gender:
            draw_button(start_box, 'START GAME', GREEN)
        else:
            draw_button(start_box, 'START GAME', passive_color)

        # Render text inside name box
        name_surface = input_font.render(user_name, True, WHITE)
        screen.blit(name_surface, (name_box.x + 5, name_box.y + 5))
        name_box.w = max(300, name_surface.get_width() + 10)

        pygame.display.update()
        
        
        
        
        #save game 
        def save_to_excel(outcome):
    """Saves game data to a uniquely named Excel file."""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"Ufo Game_{i}.xlsx"
    
    try:
        # Create a new workbook and select the active sheet
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Game Report"

        # Write headers
        headers = ["Player Name", "Gender", "Final Score", "Game Outcome", "Timestamp"]
        sheet.append(headers)

        # Write data
        data_row = [user_name, user_gender, score_value, outcome, timestamp]
        sheet.append(data_row)
        
        # Adjust column widths
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width

        # Save the file
        workbook.save(filename)
        print(f"Game data successfully saved to '{filename}'")

    except Exception as e:
        print(f"Error: Could not save data to Excel file. {e}")
        
        
        
        
        def save_to_excel(outcome):
    i = 1
    while os.path.exists(f"Ufo Game_{i}.xlsx"):
        i += 1
    
    filename = f"Ufo Game_{i}.xlsx"
    
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Game Report"

        # Write headers
        headers = ["Player Name", "Gender", "Final Score", "Game Outcome", "Timestamp"]
        sheet.append(headers)

        # Write data
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data_row = [user_name, user_gender, score_value, outcome, timestamp]
        sheet.append(data_row)
        
        # Adjust column widths
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width

        # Save the file
        workbook.save(filename)
        print(f"Game data successfully saved to '{filename}'")

    except Exception as e:
        print(f"Error: Could not save data to Excel file. {e}")
        
        
        
        
        
        
def game_loop():
    """The main loop where the game is played."""
    global player_x, player_y, player_x_change, bullet_state, bullet_x, bullet_y, score_value, lives
    global enemy_speed_multiplier  # Add this if you modify it in the function
    
    reset_game_state()
    mixer.music.play(-1)
    
    start_time = pygame.time.get_ticks()

    running = True
    while running:
        screen.blit(background_img, (0, 0))
        
        # --- Increase difficulty over time ---
        if (pygame.time.get_ticks() - start_time) > 60000: # Every 60 seconds
            enemy_speed_multiplier += 0.2
            start_time = pygame.time.get_ticks()

        # --- Event Handling ---
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                save_to_excel("Quit")
                pygame.quit()
                exit()
            
            if event.type == pygame.MOUSEBUTTONDOWN:
                if quit_button_rect.collidepoint(event.pos):
                    save_to_excel("Quit")
                    running = False # Exit game loop, return to home

            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_LEFT: player_x_change = -sensitivity
                if event.key == pygame.K_RIGHT: player_x_change = sensitivity
                if event.key == pygame.K_SPACE and bullet_state == "ready":
                    bullet_sound.play()
                    bullet_state = "fire"
                    bullet_x, bullet_y = player_x, player_y
            
            if event.type == pygame.KEYUP:
                if event.key in [pygame.K_LEFT, pygame.K_RIGHT]:
                    player_x_change = 0
        
        # --- Update Game Logic ---
        player_x += player_x_change
        player_x = max(0, min(player_x, SCREEN_WIDTH - 64))

        # Enemy movement and collision
        for enemy in list(enemies):
            enemy["x"] += enemy["x_change"]
            if enemy["x"] <= 0 or enemy["x"] >= SCREEN_WIDTH - 64:
                enemy["x_change"] *= -1
                enemy["y"] += 40
            
            if bullet_state == "fire" and math.sqrt((enemy["x"] - bullet_x)**2 + (enemy["y"] - bullet_y)**2) < 27:
                explosion_sound.play()
                bullet_state = "ready"
                score_value += 1
                enemies.remove(enemy)
                enemies.append(create_enemy())
            
            if enemy["y"] > 480:
                lives -= 1
                enemies.remove(enemy)
                enemies.append(create_enemy())
                if lives <= 0:
                    running = False # Game Over

        if bullet_state == "fire":
            bullet_y -= bullet_y_change
            if bullet_y <= 0: bullet_state = "ready"

        # --- Drawing / Rendering ---
        if bullet_state == "fire":
            screen.blit(bullet_img, (bullet_x + 16, bullet_y + 10))
        
        screen.blit(player_img, (player_x, player_y))
        
        for enemy in enemies:
            screen.blit(enemy_img, (enemy['x'], enemy['y']))
        
        show_game_ui()
        pygame.display.update()

    # --- Post-Game Over ---
    if lives <= 0:
        save_to_excel("Game Over")
        game_over_screen()
        pygame.display.update()
        
        # Wait for user to press Enter
        waiting = True
        while waiting:
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    pygame.quit()
                    exit()
                if event.type == pygame.KEYDOWN and event.key == pygame.K_RETURN:
                    waiting = False