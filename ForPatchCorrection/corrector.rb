require 'rubyXL'
require 'set'

if ARGV.length < 1
  puts "Использование: ruby process_file.rb <файл.xlsx>"
  exit
end

xlsx_file = ARGV[0]
names_file = "names_list.txt"
output_file = "Lavanda_edited.xlsx"  # Сохраняем в НОВЫЙ файл

puts "Читаю xlsx..."
workbook = RubyXL::Parser.parse(xlsx_file)

# --- Функция обработки текста ---
def process_text(text)
  return text if text.nil? || text.empty?
  
  original = text.dup
  
    # 3. …… → ...
  text.gsub!(/……/, '...')
  
  # 4. … → ...
  text.gsub!(/…/, '...')
  
  # 1. Обработка реплик в скобках「」
  text = text.gsub(/「([^」]+)」/) do |match|
    inner = $1.strip
    if inner && !inner.empty? && !inner.match?(/[?!。.]\z/)
      "「#{inner}.」"
    else
      match
    end
  end
  
  # 9. !! → !!! (только если ровно два, не больше)
  
  # 2. !! → !!! (только если ровно два)
  text.gsub!(/(?<!\!)!!(?!\!)/, '!!!')
    
  text.gsub!(/!..!!/, '!!!..')
  
  text.gsub!(/!!!..!/, '!!!..')
  
  # 5. !... → !..
  text.gsub!(/!\.{3}/, '!..')
  
  # 6. ...! → !..
  text.gsub!(/\.{3}!/, '!..')
  
  # 7. — → –
  text.gsub!(/—/, '–')
  
  # 8. Удаляем。
  text.gsub!(/。/, '')
  
  # 9. ...? → ?..
  text.gsub!(/\.{3}\?/, '?..')
  
  text
end

# --- Функция извлечения имён ---
def extract_names(text, names_set)
  return if text.nil? || text.empty?
  text.scan(/【[^】]+】/) { |match| names_set << match }
end

# --- Основные переменные ---
names = Set.new
modified_count = 0
processed_rows = 0

# --- Обработка всех листов ---
workbook.worksheets.each do |sheet|
  puts "Обрабатываю лист: #{sheet.sheet_name}"
  
  sheet.each_with_index do |row, i|
    next if row.nil? || i == 0
    
    [3, 4].each do |col_idx|
      cell = row[col_idx]
      next if cell.nil?
      
      original_text = cell.value&.to_s
      next if original_text.nil? || original_text.empty?
      
      processed_rows += 1
      extract_names(original_text, names)
      
      processed_text = process_text(original_text)
      
      if processed_text != original_text
        # Для rubyXL 3.4.35 - модифицируем через worksheet.add_cell
        begin
          sheet.add_cell(i, col_idx, processed_text)
        rescue
          # Если не работает, пробуем прямой доступ
          begin
            cell.instance_variable_set(:@value, processed_text)
          rescue
            puts "  ⚠ Не удалось записать строку #{i + 1}"
          end
        end
        
        puts "  ✓ Строка #{i + 1}, #{col_idx == 3 ? 'TL' : 'TLC'}:"
        puts "    Было: #{original_text}"
        puts "    Стало: #{processed_text}"
        puts
        modified_count += 1
      end
    end
  end
end

# --- Сохранение в НОВЫЙ файл ---
puts "Сохраняю изменения в #{output_file}..."
workbook.write(output_file)

File.open(names_file, "w:utf-8") do |f|
  names.to_a.sort.each { |name| f.puts name }
end

puts "\n" + "=" * 50
puts "ГОТОВО!"
puts "=" * 50
puts "Всего обработано ячеек: #{processed_rows}"
puts "Изменено строк: #{modified_count}"
puts "Найдено имён: #{names.size}"
puts "Оригинал: #{xlsx_file}"
puts "Результат: #{output_file}"
puts "Список имён: #{names_file}"