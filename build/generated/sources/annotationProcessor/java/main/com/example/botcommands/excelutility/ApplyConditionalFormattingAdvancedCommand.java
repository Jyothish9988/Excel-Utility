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

public final class ApplyConditionalFormattingAdvancedCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(ApplyConditionalFormattingAdvancedCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    ApplyConditionalFormattingAdvanced command = new ApplyConditionalFormattingAdvanced();
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

    if(parameters.containsKey("greaterThan") && parameters.get("greaterThan") != null && parameters.get("greaterThan").get() != null) {
      convertedParameters.put("greaterThan", parameters.get("greaterThan").get());
      if(convertedParameters.get("greaterThan") !=null && !(convertedParameters.get("greaterThan") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","greaterThan", "Boolean", parameters.get("greaterThan").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("greaterThanValue") && parameters.get("greaterThanValue") != null && parameters.get("greaterThanValue").get() != null) {
      convertedParameters.put("greaterThanValue", parameters.get("greaterThanValue").get());
      if(convertedParameters.get("greaterThanValue") !=null && !(convertedParameters.get("greaterThanValue") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","greaterThanValue", "String", parameters.get("greaterThanValue").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("lessThan") && parameters.get("lessThan") != null && parameters.get("lessThan").get() != null) {
      convertedParameters.put("lessThan", parameters.get("lessThan").get());
      if(convertedParameters.get("lessThan") !=null && !(convertedParameters.get("lessThan") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","lessThan", "Boolean", parameters.get("lessThan").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("lessThanValue") && parameters.get("lessThanValue") != null && parameters.get("lessThanValue").get() != null) {
      convertedParameters.put("lessThanValue", parameters.get("lessThanValue").get());
      if(convertedParameters.get("lessThanValue") !=null && !(convertedParameters.get("lessThanValue") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","lessThanValue", "String", parameters.get("lessThanValue").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("between") && parameters.get("between") != null && parameters.get("between").get() != null) {
      convertedParameters.put("between", parameters.get("between").get());
      if(convertedParameters.get("between") !=null && !(convertedParameters.get("between") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","between", "Boolean", parameters.get("between").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("betweenValue1") && parameters.get("betweenValue1") != null && parameters.get("betweenValue1").get() != null) {
      convertedParameters.put("betweenValue1", parameters.get("betweenValue1").get());
      if(convertedParameters.get("betweenValue1") !=null && !(convertedParameters.get("betweenValue1") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","betweenValue1", "String", parameters.get("betweenValue1").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("betweenValue2") && parameters.get("betweenValue2") != null && parameters.get("betweenValue2").get() != null) {
      convertedParameters.put("betweenValue2", parameters.get("betweenValue2").get());
      if(convertedParameters.get("betweenValue2") !=null && !(convertedParameters.get("betweenValue2") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","betweenValue2", "String", parameters.get("betweenValue2").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("equalTo") && parameters.get("equalTo") != null && parameters.get("equalTo").get() != null) {
      convertedParameters.put("equalTo", parameters.get("equalTo").get());
      if(convertedParameters.get("equalTo") !=null && !(convertedParameters.get("equalTo") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","equalTo", "Boolean", parameters.get("equalTo").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("equalToValue") && parameters.get("equalToValue") != null && parameters.get("equalToValue").get() != null) {
      convertedParameters.put("equalToValue", parameters.get("equalToValue").get());
      if(convertedParameters.get("equalToValue") !=null && !(convertedParameters.get("equalToValue") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","equalToValue", "String", parameters.get("equalToValue").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("blank") && parameters.get("blank") != null && parameters.get("blank").get() != null) {
      convertedParameters.put("blank", parameters.get("blank").get());
      if(convertedParameters.get("blank") !=null && !(convertedParameters.get("blank") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","blank", "Boolean", parameters.get("blank").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("textContains") && parameters.get("textContains") != null && parameters.get("textContains").get() != null) {
      convertedParameters.put("textContains", parameters.get("textContains").get());
      if(convertedParameters.get("textContains") !=null && !(convertedParameters.get("textContains") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textContains", "Boolean", parameters.get("textContains").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("textValue") && parameters.get("textValue") != null && parameters.get("textValue").get() != null) {
      convertedParameters.put("textValue", parameters.get("textValue").get());
      if(convertedParameters.get("textValue") !=null && !(convertedParameters.get("textValue") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textValue", "String", parameters.get("textValue").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("duplicates") && parameters.get("duplicates") != null && parameters.get("duplicates").get() != null) {
      convertedParameters.put("duplicates", parameters.get("duplicates").get());
      if(convertedParameters.get("duplicates") !=null && !(convertedParameters.get("duplicates") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","duplicates", "Boolean", parameters.get("duplicates").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("fillColorHex") && parameters.get("fillColorHex") != null && parameters.get("fillColorHex").get() != null) {
      convertedParameters.put("fillColorHex", parameters.get("fillColorHex").get());
      if(convertedParameters.get("fillColorHex") !=null && !(convertedParameters.get("fillColorHex") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","fillColorHex", "String", parameters.get("fillColorHex").get().getClass().getSimpleName()));
      }
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("filePath"),(String)convertedParameters.get("sheetName"),(String)convertedParameters.get("columnList"),(Boolean)convertedParameters.get("greaterThan"),(String)convertedParameters.get("greaterThanValue"),(Boolean)convertedParameters.get("lessThan"),(String)convertedParameters.get("lessThanValue"),(Boolean)convertedParameters.get("between"),(String)convertedParameters.get("betweenValue1"),(String)convertedParameters.get("betweenValue2"),(Boolean)convertedParameters.get("equalTo"),(String)convertedParameters.get("equalToValue"),(Boolean)convertedParameters.get("blank"),(Boolean)convertedParameters.get("textContains"),(String)convertedParameters.get("textValue"),(Boolean)convertedParameters.get("duplicates"),(String)convertedParameters.get("fillColorHex")));
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
