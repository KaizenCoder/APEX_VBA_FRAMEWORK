#!/usr/bin/env ruby
# encoding: UTF-8
# ============================================================================
# Script de Validation Post-Refactoring APEX Framework
# Référence: APEX-VAL-RUBY-001
# Date: 2025-04-30
# ============================================================================

require 'win32ole'
require 'json'
require 'time'
require 'fileutils'

class ApexValidationRunner
  attr_reader :results, :excel, :workbook

  def initialize(target_file, output_format = 'markdown')
    if target_file.start_with?('Avis_interoperabilite/')
      @target_file = File.expand_path(target_file)
    else
      @target_file = File.expand_path(File.join(Dir.pwd, target_file))
    end
    
    @output_format = output_format
    @results = {
      metadata: {
        timestamp: Time.now.iso8601,
        target_file: @target_file,
        reference: "APEX-VAL-RUBY-001"
      },
      tests: [],
      summary: {
        total: 0,
        passed: 0,
        failed: 0,
        errors: 0
      }
    }
    @excel = nil
    @workbook = nil
  end

  def run_all_tests
    open_excel
    
    begin
      # Vérifier si le module VBA est présent dans le classeur
      unless module_exists?("TestRealExcelInterop")
        puts "\n⚠️ ATTENTION: Le module 'TestRealExcelInterop' n'a pas été trouvé dans le classeur."
        puts "Veuillez importer le module VBA en suivant les instructions dans README_VALIDATION.md."
        puts "1. Ouvrir l'éditeur VBA (Alt+F11)"
        puts "2. Importer le fichier 'TestRealExcelInterop.bas' via le menu 'Fichier > Importer un fichier'"
        puts "3. Relancer ce script après avoir importé le module\n"
        
        add_result(
          "Module_VBA_Manquant", 
          "Vérification de la présence du module VBA de test", 
          false, 
          "Le module 'TestRealExcelInterop' doit être importé dans le classeur"
        )
        return results
      end
      
      # Tests dans l'ordre de dépendance
      test_appcontext_initialization
      test_excel_factory_access
      test_logger_access
      test_cache_minimal
      
      # Tests supplémentaires spécifiques si nécessaire
      # ...
    rescue => e
      add_result(
        "Exception_Globale", 
        "Une erreur non gérée s'est produite pendant l'exécution", 
        false, 
        e.message
      )
    ensure
      close_excel
    end
    
    generate_summary
    save_report
    
    results
  end

  private
  
  def module_exists?(module_name)
    begin
      # Essayer d'accéder au module VBA
      @excel.run("'#{module_name}.TestIAppContextInitialization")
      return true
    rescue => e
      # Vérifie le message d'erreur pour distinguer si c'est un problème d'accès 
      # ou si le module n'existe pas
      if e.message.include?("Impossible d'exécuter la macro") || 
         e.message.include?("Can't run the macro")
        return false
      end
      # Si l'erreur est autre, retourne true car le module existe probablement
      # mais a une autre erreur
      return true
    end
  end
  
  def open_excel
    begin
      puts "Ouverture d'Excel et du fichier cible: #{@target_file}"
      @excel = WIN32OLE.new('Excel.Application')
      @excel.visible = true # À des fins de débogage, mettre à false en production
      
      # Vérifier si le fichier existe
      unless File.exist?(@target_file)
        raise "Le fichier cible n'existe pas: #{@target_file}"
      end
      
      # Ouvrir le classeur
      @workbook = @excel.Workbooks.Open(@target_file)
      add_result("Ouverture_Excel", "Ouverture du fichier Excel cible", true)
      
    rescue => e
      add_result("Ouverture_Excel", "Ouverture du fichier Excel cible", false, e.message)
      raise e # Remonter l'erreur pour arrêter les tests
    end
  end
  
  def close_excel
    begin
      @workbook.Close(false) if @workbook && @workbook.respond_to?(:Close) # Ne pas sauvegarder
      @excel.Quit if @excel
      @excel = nil
      @workbook = nil
      add_result("Fermeture_Excel", "Fermeture propre d'Excel", true)
    rescue => e
      add_result("Fermeture_Excel", "Fermeture propre d'Excel", false, e.message)
    end
  end

  def add_result(test_name, description, passed, error_message = nil)
    # S'assurer que l'erreur est encodée correctement
    error_message = error_message.to_s.encode('UTF-8', invalid: :replace, undef: :replace, replace: '?') if error_message
    
    @results[:tests] << {
      name: test_name,
      description: description,
      passed: passed,
      error: error_message,
      timestamp: Time.now.iso8601
    }
    
    # Mettre à jour le compteur approprié
    @results[:summary][:total] += 1
    if passed
      @results[:summary][:passed] += 1
    elsif error_message
      @results[:summary][:errors] += 1
    else
      @results[:summary][:failed] += 1
    end
    
    # Afficher le résultat dans la console
    status = passed ? "✅ RÉUSSI" : "❌ ÉCHEC"
    puts "#{status} - #{test_name}: #{description}"
    puts "   Erreur: #{error_message}" if error_message
  end

  def test_appcontext_initialization
    begin
      # Exécuter la fonction VBA qui initialise et retourne un statut
      result = @excel.run("TestIAppContextInitialization")
      
      # Analyser le résultat
      success = (result.to_s.downcase == "true")
      error_msg = success ? nil : "L'initialisation d'IAppContext a échoué"
      
      add_result(
        "IAppContext_Initialization", 
        "Vérification de l'initialisation correcte d'IAppContext", 
        success, 
        error_msg
      )
    rescue => e
      add_result(
        "IAppContext_Initialization", 
        "Vérification de l'initialisation correcte d'IAppContext", 
        false, 
        "Exception: #{e.message}"
      )
    end
  end
  
  def test_excel_factory_access
    begin
      # Exécuter la fonction VBA qui teste l'accès à ExcelFactory via IAppContext
      result = @excel.run("TestExcelFactoryAccess")
      
      # Analyser le résultat
      success = (result.to_s.downcase == "true")
      error_msg = success ? nil : "L'accès à ExcelFactory via IAppContext a échoué"
      
      add_result(
        "ExcelFactory_Access", 
        "Vérification de l'accès à ExcelFactory via IAppContext", 
        success, 
        error_msg
      )
    rescue => e
      add_result(
        "ExcelFactory_Access", 
        "Vérification de l'accès à ExcelFactory via IAppContext", 
        false, 
        "Exception: #{e.message}"
      )
    end
  end
  
  def test_logger_access
    begin
      # Exécuter la fonction VBA qui teste l'accès au Logger via IAppContext
      result = @excel.run("TestLoggerAccess")
      
      # Analyser le résultat
      success = (result.to_s.downcase == "true")
      error_msg = success ? nil : "L'accès au Logger via IAppContext a échoué"
      
      add_result(
        "Logger_Access", 
        "Vérification de l'accès au Logger via IAppContext", 
        success, 
        error_msg
      )
    rescue => e
      add_result(
        "Logger_Access", 
        "Vérification de l'accès au Logger via IAppContext", 
        false, 
        "Exception: #{e.message}"
      )
    end
  end
  
  def test_cache_minimal
    begin
      # Exécuter la fonction VBA qui effectue un test minimal du cache
      result = @excel.run("TestCacheMinimal")
      
      # Analyser le résultat
      success = (result.to_s.downcase == "true")
      error_msg = success ? nil : "Le test minimal du cache a échoué"
      
      add_result(
        "Cache_Minimal", 
        "Vérification du fonctionnement minimal du cache", 
        success, 
        error_msg
      )
    rescue => e
      add_result(
        "Cache_Minimal", 
        "Vérification du fonctionnement minimal du cache", 
        false, 
        "Exception: #{e.message}"
      )
    end
  end
  
  def generate_summary
    @results[:summary][:success_rate] = (@results[:summary][:passed].to_f / @results[:summary][:total] * 100).round(2)
    @results[:summary][:end_time] = Time.now.iso8601
    @results[:summary][:duration] = Time.parse(@results[:summary][:end_time]) - Time.parse(@results[:metadata][:timestamp])
    
    puts "\n=== RÉSUMÉ DES TESTS ==="
    puts "Total: #{@results[:summary][:total]}"
    puts "Réussis: #{@results[:summary][:passed]}"
    puts "Échecs: #{@results[:summary][:failed]}"
    puts "Erreurs: #{@results[:summary][:errors]}"
    puts "Taux de réussite: #{@results[:summary][:success_rate]}%"
    puts "Durée: #{@results[:summary][:duration]} secondes"
  end
  
  def save_report
    begin
      # Créer le dossier de rapports si nécessaire
      report_dir = File.join(File.dirname(@target_file), 'reports')
      FileUtils.mkdir_p(report_dir) unless Dir.exist?(report_dir)
      
      # Générer le nom du fichier avec timestamp
      timestamp = Time.now.strftime('%Y%m%d_%H%M%S')
      
      if @output_format.downcase == 'json'
        # Sauvegarder en JSON
        filename = File.join(report_dir, "validation_report_#{timestamp}.json")
        File.write(filename, JSON.pretty_generate(@results))
      else
        # Sauvegarder en Markdown
        filename = File.join(report_dir, "validation_report_#{timestamp}.md")
        File.write(filename, generate_markdown_report)
      end
      
      puts "Rapport sauvegardé: #{filename}"
    rescue => e
      puts "Erreur lors de la sauvegarde du rapport: #{e.message}"
      puts e.backtrace.join("\n")
    end
  end
  
  def generate_markdown_report
    report = []
    
    # En-tête
    report << "# Rapport de Validation Post-Refactoring APEX Framework"
    report << ""
    report << "**Référence:** #{@results[:metadata][:reference]}"
    report << "**Date d'exécution:** #{Time.parse(@results[:metadata][:timestamp]).strftime('%Y-%m-%d %H:%M:%S')}"
    report << "**Fichier cible:** #{@results[:metadata][:target_file]}"
    report << ""
    
    # Résumé
    report << "## Résumé"
    report << ""
    report << "| Métrique | Valeur |"
    report << "| -------- | ------ |"
    report << "| Tests totaux | #{@results[:summary][:total]} |"
    report << "| Tests réussis | #{@results[:summary][:passed]} |"
    report << "| Tests échoués | #{@results[:summary][:failed]} |"
    report << "| Erreurs | #{@results[:summary][:errors]} |"
    report << "| Taux de réussite | #{@results[:summary][:success_rate]}% |"
    report << "| Durée d'exécution | #{@results[:summary][:duration]} secondes |"
    report << ""
    
    # Résultats détaillés
    report << "## Résultats détaillés"
    report << ""
    
    @results[:tests].each do |test|
      status = test[:passed] ? "✅ RÉUSSI" : "❌ ÉCHEC"
      report << "### #{status} - #{test[:name]}"
      report << ""
      report << "**Description:** #{test[:description]}"
      report << "**Exécuté à:** #{Time.parse(test[:timestamp]).strftime('%H:%M:%S')}"
      
      if !test[:passed] && test[:error]
        report << ""
        report << "**Détail de l'erreur:**"
        report << "```"
        report << test[:error]
        report << "```"
      end
      
      report << ""
    end
    
    # Conclusion
    if @results[:summary][:failed] == 0 && @results[:summary][:errors] == 0
      report << "## Conclusion"
      report << ""
      report << "✅ **VALIDATION RÉUSSIE** - Tous les tests ont passé avec succès."
    else
      report << "## Conclusion"
      report << ""
      report << "❌ **VALIDATION ÉCHOUÉE** - Des problèmes ont été détectés. Voir les détails ci-dessus."
    end
    
    report.join("\n")
  end
end

# Point d'entrée principal
if __FILE__ == $0
  puts "========================================================================"
  puts "  Script de Validation Post-Refactoring APEX Framework"
  puts "  Référence: APEX-VAL-RUBY-001"
  puts "  Date: 2025-04-30"
  puts "========================================================================"
  puts ""
  puts "IMPORTANT: Avant d'exécuter ce script, assurez-vous que:"
  puts "1. Le module VBA 'TestRealExcelInterop.bas' a été importé dans le classeur cible"
  puts "2. Excel est installé et les macros sont activées"
  puts "3. Le classeur cible existe à l'emplacement spécifié"
  puts ""
  
  # Paramètres
  target_file = ARGV[0] || 'TestExcelInterop.xlsm'
  output_format = ARGV[1] || 'markdown'
  
  # Validation des paramètres
  unless ['json', 'markdown'].include?(output_format.downcase)
    puts "Format de sortie non valide. Utilisation de 'markdown' par défaut."
    output_format = 'markdown'
  end
  
  # Exécution des tests
  runner = ApexValidationRunner.new(target_file, output_format)
  runner.run_all_tests
end 