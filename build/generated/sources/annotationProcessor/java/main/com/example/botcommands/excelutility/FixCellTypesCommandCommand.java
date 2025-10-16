package com.example.botcommands.excelutility;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import java.lang.Boolean;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class FixCellTypesCommandCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(FixCellTypesCommandCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    FixCellTypesCommand command = new FixCellTypesCommand();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("filePath") && parameters.get("filePath") != null && parameters.get("filePath").get() != null) {
      convertedParameters.put("filePath", parameters.get("filePath").get());
      if(convertedParameters.get("filePath") !=null && !(convertedParameters.get("filePath") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","filePath", "String", parameters.get("filePath").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("sheetName") && parameters.get("sheetName") != null && parameters.get("sheetName").get() != null) {
      convertedParameters.put("sheetName", parameters.get("sheetName").get());
      if(convertedParameters.get("sheetName") !=null && !(convertedParameters.get("sheetName") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetName", "String", parameters.get("sheetName").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("columnList") && parameters.get("columnList") != null && parameters.get("columnList").get() != null) {
      convertedParameters.put("columnList", parameters.get("columnList").get());
      if(convertedParameters.get("columnList") !=null && !(convertedParameters.get("columnList") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","columnList", "String", parameters.get("columnList").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("fixNumbers") && parameters.get("fixNumbers") != null && parameters.get("fixNumbers").get() != null) {
      convertedParameters.put("fixNumbers", parameters.get("fixNumbers").get());
      if(convertedParameters.get("fixNumbers") !=null && !(convertedParameters.get("fixNumbers") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","fixNumbers", "Boolean", parameters.get("fixNumbers").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("fixDates") && parameters.get("fixDates") != null && parameters.get("fixDates").get() != null) {
      convertedParameters.put("fixDates", parameters.get("fixDates").get());
      if(convertedParameters.get("fixDates") !=null && !(convertedParameters.get("fixDates") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","fixDates", "Boolean", parameters.get("fixDates").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("outputFile") && parameters.get("outputFile") != null && parameters.get("outputFile").get() != null) {
      convertedParameters.put("outputFile", parameters.get("outputFile").get());
      if(convertedParameters.get("outputFile") !=null && !(convertedParameters.get("outputFile") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","outputFile", "String", parameters.get("outputFile").get().getClass().getSimpleName()));
      }
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("filePath"),(String)convertedParameters.get("sheetName"),(String)convertedParameters.get("columnList"),(Boolean)convertedParameters.get("fixNumbers"),(Boolean)convertedParameters.get("fixDates"),(String)convertedParameters.get("outputFile")));
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","action"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }

  public Map<String, Value> executeAndReturnMany(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return null;
  }
}
